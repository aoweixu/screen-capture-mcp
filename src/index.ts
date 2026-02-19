#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import sharp from "sharp";
import { execSync } from "child_process";

const server = new McpServer({
  name: "screen-capture-mcp",
  version: "1.0.0",
});

function runPowerShell(script: string): Buffer {
  const encoded = Buffer.from(script, "utf16le").toString("base64");
  const result = execSync(
    `powershell -NoProfile -NonInteractive -EncodedCommand ${encoded}`,
    { encoding: "utf-8", maxBuffer: 50 * 1024 * 1024, timeout: 15000 }
  );
  return Buffer.from(result.trim(), "base64");
}

interface ScreenInfo {
  x: number;
  y: number;
  width: number;
  height: number;
}

function getScreenLayout(): ScreenInfo[] {
  const result = execSync(
    `powershell -NoProfile -NonInteractive -Command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Screen]::AllScreens | ForEach-Object { Write-Output ('' + $_.Bounds.X + ',' + $_.Bounds.Y + ',' + $_.Bounds.Width + ',' + $_.Bounds.Height) }"`,
    { encoding: "utf-8", timeout: 10000 }
  );
  return result.trim().split(/\r?\n/).map((line) => {
    const [x, y, width, height] = line.split(",").map(Number);
    return { x, y, width, height };
  });
}

function captureSingleScreen(screen: ScreenInfo): Buffer {
  return runPowerShell(`
Add-Type -AssemblyName System.Drawing

$bmp = New-Object System.Drawing.Bitmap(${screen.width}, ${screen.height})
$graphics = [System.Drawing.Graphics]::FromImage($bmp)
$graphics.CopyFromScreen(${screen.x}, ${screen.y}, 0, 0, (New-Object System.Drawing.Size(${screen.width}, ${screen.height})))
$graphics.Dispose()

$ms = New-Object System.IO.MemoryStream
$bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

[Convert]::ToBase64String($ms.ToArray())
$ms.Dispose()
`);
}

async function captureAllScreens(): Promise<Buffer> {
  const screens = getScreenLayout();
  if (screens.length === 1) {
    return captureSingleScreen(screens[0]);
  }

  const minX = Math.min(...screens.map((s) => s.x));
  const minY = Math.min(...screens.map((s) => s.y));
  const maxX = Math.max(...screens.map((s) => s.x + s.width));
  const maxY = Math.max(...screens.map((s) => s.y + s.height));
  const totalWidth = maxX - minX;
  const totalHeight = maxY - minY;

  const captures = screens.map((screen) => ({
    input: captureSingleScreen(screen),
    left: screen.x - minX,
    top: screen.y - minY,
  }));

  return sharp({
    create: {
      width: totalWidth,
      height: totalHeight,
      channels: 4,
      background: { r: 0, g: 0, b: 0, alpha: 1 },
    },
  })
    .composite(
      captures.map((c) => ({
        input: c.input,
        left: c.left,
        top: c.top,
      }))
    )
    .png()
    .toBuffer();
}

function captureWindowByTitle(title: string): Buffer {
  const safeTitle = title.replace(/'/g, "''");

  return runPowerShell(`
Add-Type -AssemblyName System.Drawing

Add-Type @'
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
}
public struct RECT {
    public int Left; public int Top; public int Right; public int Bottom;
}
'@

$proc = Get-Process | Where-Object { $_.MainWindowTitle -like '*${safeTitle}*' -and $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if (-not $proc) { Write-Error "Window not found: ${safeTitle}"; exit 1 }

$hwnd = $proc.MainWindowHandle
$rect = New-Object RECT
[Win32]::GetWindowRect($hwnd, [ref]$rect) | Out-Null

$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top

if ($width -le 0 -or $height -le 0) { Write-Error "Invalid window dimensions"; exit 1 }

$bmp = New-Object System.Drawing.Bitmap($width, $height)
$graphics = [System.Drawing.Graphics]::FromImage($bmp)
$graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, (New-Object System.Drawing.Size($width, $height)))
$graphics.Dispose()

$ms = New-Object System.IO.MemoryStream
$bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

[Convert]::ToBase64String($ms.ToArray())
$ms.Dispose()
`);
}

async function resizeImage(
  pngBuffer: Buffer,
  targetWidth = 2560
): Promise<Buffer> {
  const metadata = await sharp(pngBuffer).metadata();
  if (metadata.width && metadata.width > targetWidth) {
    return sharp(pngBuffer)
      .resize({ width: targetWidth, withoutEnlargement: true })
      .png()
      .toBuffer();
  }
  return pngBuffer;
}

server.tool(
  "take_screenshot",
  "Captures a screenshot of all displays or a specific window. Returns the image as a PNG. If window_title is provided, captures only that window (partial title match). Otherwise captures all monitors combined.",
  {
    window_title: z
      .string()
      .optional()
      .describe(
        "Optional window title to capture (partial match). If omitted, captures all monitors."
      ),
  },
  async ({ window_title }) => {
    try {
      let pngBuffer: Buffer;

      if (window_title) {
        pngBuffer = captureWindowByTitle(window_title);
      } else {
        pngBuffer = await captureAllScreens();
      }

      pngBuffer = await resizeImage(pngBuffer);
      const base64 = pngBuffer.toString("base64");

      return {
        content: [
          {
            type: "image" as const,
            data: base64,
            mimeType: "image/png",
          },
        ],
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        content: [
          { type: "text" as const, text: `Screenshot failed: ${message}` },
        ],
        isError: true,
      };
    }
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error("Server failed to start:", error);
  process.exit(1);
});
