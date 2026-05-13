// Generates solid-color placeholder PNGs for the extension icons.
// Uses only Node.js built-ins — no external dependencies required.
// Icon color matches the accent: #4f6ef7 = rgb(79, 110, 247)

const zlib = require('zlib')
const fs   = require('fs')
const path = require('path')

const CRC_TABLE = (() => {
  const t = new Uint32Array(256)
  for (let i = 0; i < 256; i++) {
    let c = i
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[i] = c
  }
  return t
})()

function crc32(data) {
  let crc = 0xffffffff
  for (let i = 0; i < data.length; i++) crc = CRC_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

function u32be(n) {
  const b = Buffer.alloc(4)
  b.writeUInt32BE(n, 0)
  return b
}

function pngChunk(type, data) {
  const typeBytes = Buffer.from(type, 'ascii')
  const crcInput  = Buffer.concat([typeBytes, data])
  return Buffer.concat([u32be(data.length), typeBytes, data, u32be(crc32(crcInput))])
}

function createSolidPNG(size, r, g, b) {
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])

  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(size, 0)
  ihdr.writeUInt32BE(size, 4)
  ihdr[8]  = 8  // bit depth
  ihdr[9]  = 2  // color type: RGB
  ihdr[10] = 0  // compression method
  ihdr[11] = 0  // filter method
  ihdr[12] = 0  // interlace method

  // Row format: filter byte (0 = None) + R G B per pixel
  const raw = Buffer.alloc(size * (1 + size * 3))
  for (let y = 0; y < size; y++) {
    const rowBase = y * (1 + size * 3)
    raw[rowBase] = 0
    for (let x = 0; x < size; x++) {
      const px = rowBase + 1 + x * 3
      raw[px] = r; raw[px + 1] = g; raw[px + 2] = b
    }
  }

  const idat = zlib.deflateSync(raw, { level: 9 })

  return Buffer.concat([
    sig,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', idat),
    pngChunk('IEND', Buffer.alloc(0)),
  ])
}

const outDir = path.join(__dirname, '..', 'public', 'icons')
fs.mkdirSync(outDir, { recursive: true })

// Accent color #4f6ef7
const [R, G, B] = [79, 110, 247]

for (const size of [16, 32, 48, 128]) {
  const png  = createSolidPNG(size, R, G, B)
  const dest = path.join(outDir, `icon${size}.png`)
  fs.writeFileSync(dest, png)
  console.log(`  ✓ icon${size}.png  (${png.length} bytes)`)
}

console.log('\nIcons written to public/icons/')
