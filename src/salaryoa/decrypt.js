// Decrypt password-protected Office files (ECMA-376 Agile encryption) in the
// browser using WebCrypto. Mirrors msoffcrypto-tool's agile path, which covers
// every APAC file Excel produces with "Encrypt with Password".
import * as cfbModule from "cfb";

const CFB = cfbModule.read ? cfbModule : cfbModule.default;

const subtle = globalThis.crypto.subtle;

function concat(...arrays) {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const a of arrays) {
    out.set(a, off);
    off += a.length;
  }
  return out;
}

function le32(n) {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n, true);
  return b;
}

function b64(s) {
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function utf16le(s) {
  const out = new Uint8Array(s.length * 2);
  const dv = new DataView(out.buffer);
  for (let i = 0; i < s.length; i++) dv.setUint16(i * 2, s.charCodeAt(i), true);
  return out;
}

async function hash(algo, data) {
  return new Uint8Array(await subtle.digest(algo, data));
}

// AES-CBC decrypt WITHOUT PKCS#7 padding: WebCrypto insists on padding, so we
// append one encrypted full-pad block to make the unpad succeed.
async function aesCbcDecryptNoPad(keyBytes, iv, ciphertext) {
  const key = await subtle.importKey("raw", keyBytes, { name: "AES-CBC" }, false, ["encrypt", "decrypt"]);
  const lastBlock = ciphertext.slice(ciphertext.length - 16);
  const padBlock = new Uint8Array(await subtle.encrypt({ name: "AES-CBC", iv: lastBlock }, key, new Uint8Array(0)));
  const padded = concat(ciphertext, padBlock);
  const plain = new Uint8Array(await subtle.decrypt({ name: "AES-CBC", iv }, key, padded));
  return plain.slice(0, ciphertext.length);
}

function attr(xml, tag, name) {
  // Pull attribute `name` from the first <...tag ...> element (namespace-blind).
  const m = xml.match(new RegExp(`<[^>]*${tag}[^>]*?\\s${name}="([^"]*)"`));
  return m ? m[1] : null;
}

function fitIv(saltBytes, blockSize) {
  if (saltBytes.length === blockSize) return saltBytes;
  if (saltBytes.length > blockSize) return saltBytes.slice(0, blockSize);
  const out = new Uint8Array(blockSize).fill(0x36);
  out.set(saltBytes);
  return out;
}

const BLOCK_VERIFIER_INPUT = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLOCK_VERIFIER_VALUE = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLOCK_KEY_VALUE = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);

const HASH_NAMES = { SHA1: "SHA-1", SHA256: "SHA-256", SHA384: "SHA-384", SHA512: "SHA-512" };

export function isEncrypted(arrayBuffer) {
  const head = new Uint8Array(arrayBuffer.slice(0, 8));
  // ZIP (plain xlsx) starts with PK; CFB (encrypted container) with D0 CF 11 E0
  return head[0] === 0xd0 && head[1] === 0xcf;
}

export async function decryptOffice(arrayBuffer, password) {
  const cfb = CFB.read(new Uint8Array(arrayBuffer), { type: "buffer" });
  const info = CFB.find(cfb, "EncryptionInfo");
  const pkg = CFB.find(cfb, "EncryptedPackage");
  if (!info || !pkg) throw new Error("File is not a recognised encrypted Office file.");

  const infoBytes = new Uint8Array(info.content);
  const verMajor = new DataView(infoBytes.buffer, infoBytes.byteOffset).getUint16(0, true);
  if (verMajor !== 4) {
    throw new Error("Unsupported (legacy) encryption format — open in Excel and re-save with password, or remove the password.");
  }
  const xml = new TextDecoder("utf-8").decode(infoBytes.slice(8));

  // keyData (package encryption parameters)
  const kdSalt = b64(attr(xml, "keyData", "saltValue"));
  const kdBlockSize = parseInt(attr(xml, "keyData", "blockSize"), 10);
  const kdHash = HASH_NAMES[attr(xml, "keyData", "hashAlgorithm")] || "SHA-512";
  const kdKeyBits = parseInt(attr(xml, "keyData", "keyBits"), 10);

  // encryptedKey (password key encryptor)
  const ekSalt = b64(attr(xml, "encryptedKey", "saltValue"));
  const ekSpin = parseInt(attr(xml, "encryptedKey", "spinCount"), 10);
  const ekHash = HASH_NAMES[attr(xml, "encryptedKey", "hashAlgorithm")] || "SHA-512";
  const ekKeyBits = parseInt(attr(xml, "encryptedKey", "keyBits"), 10);
  const ekBlockSize = parseInt(attr(xml, "encryptedKey", "blockSize"), 10) || 16;
  const encVerifierInput = b64(attr(xml, "encryptedKey", "encryptedVerifierHashInput"));
  const encVerifierValue = b64(attr(xml, "encryptedKey", "encryptedVerifierHashValue"));
  const encKeyValue = b64(attr(xml, "encryptedKey", "encryptedKeyValue"));

  // Iterated password hash: h = H(salt || pwd); h = H(LE32(i) || h) × spinCount
  let h = await hash(ekHash, concat(ekSalt, utf16le(password)));
  for (let i = 0; i < ekSpin; i++) h = await hash(ekHash, concat(le32(i), h));

  const deriveKey = async (blockKey, bits) => (await hash(ekHash, concat(h, blockKey))).slice(0, bits / 8);
  const ekIv = fitIv(ekSalt, ekBlockSize);

  // Verify password
  const verifierInput = await aesCbcDecryptNoPad(await deriveKey(BLOCK_VERIFIER_INPUT, ekKeyBits), ekIv, encVerifierInput);
  const verifierHash = await hash(ekHash, verifierInput);
  const expectedHash = await aesCbcDecryptNoPad(await deriveKey(BLOCK_VERIFIER_VALUE, ekKeyBits), ekIv, encVerifierValue);
  for (let i = 0; i < verifierHash.length; i++) {
    if (verifierHash[i] !== expectedHash[i]) throw new Error("Wrong password.");
  }

  // Unwrap the package key
  const packageKey = (await aesCbcDecryptNoPad(await deriveKey(BLOCK_KEY_VALUE, ekKeyBits), ekIv, encKeyValue))
    .slice(0, kdKeyBits / 8);

  // Decrypt the package in 4096-byte segments; per-segment IV = H(kdSalt || LE32(i))
  const pkgBytes = new Uint8Array(pkg.content);
  const totalSize = Number(new DataView(pkgBytes.buffer, pkgBytes.byteOffset).getBigUint64(0, true));
  const cipher = pkgBytes.slice(8);
  const SEG = 4096;
  const chunks = [];
  for (let i = 0; i * SEG < cipher.length; i++) {
    const seg = cipher.slice(i * SEG, (i + 1) * SEG);
    const iv = fitIv((await hash(kdHash, concat(kdSalt, le32(i)))).slice(0, kdBlockSize), kdBlockSize);
    chunks.push(await aesCbcDecryptNoPad(packageKey, iv, seg));
  }
  return concat(...chunks).slice(0, totalSize).buffer;
}
