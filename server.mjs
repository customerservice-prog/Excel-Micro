import { createServer } from 'node:http'
import { readFile, stat } from 'node:fs/promises'
import { extname, join, normalize } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = fileURLToPath(new URL('./dist/', import.meta.url))
const port = Number(process.env.PORT || 4173)

const mime = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
}

function safePath(urlPath) {
  const pathname = decodeURIComponent(urlPath.split('?')[0])
  const relative = normalize(pathname).replace(/^([/\\])+/, '')
  return join(root, relative)
}

async function serveFile(pathname, response) {
  const file = await readFile(pathname)
  response.writeHead(200, {
    'Content-Type': mime[extname(pathname)] || 'application/octet-stream',
    'Cache-Control': extname(pathname) === '.html' ? 'no-cache' : 'public, max-age=31536000, immutable',
  })
  response.end(file)
}

createServer(async (request, response) => {
  try {
    const requested = safePath(request.url || '/')
    const info = await stat(requested).catch(() => null)

    if (info?.isFile()) {
      await serveFile(requested, response)
      return
    }

    if (info?.isDirectory()) {
      const index = join(requested, 'index.html')
      const indexInfo = await stat(index).catch(() => null)
      if (indexInfo?.isFile()) {
        await serveFile(index, response)
        return
      }
    }

    await serveFile(join(root, 'index.html'), response)
  } catch (error) {
    response.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' })
    response.end('Excel Micro failed to load.')
    console.error(error)
  }
}).listen(port, '0.0.0.0', () => {
  console.log(`Excel Micro listening on http://0.0.0.0:${port}`)
})
