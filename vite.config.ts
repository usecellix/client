import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import os from 'os';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Use Office Add-in certificate for HTTPS
const certPath = path.join(os.homedir(), '.office-addin-dev-certs', 'localhost.crt');
const keyPath = path.join(os.homedir(), '.office-addin-dev-certs', 'localhost.key');

let httpsConfig = undefined;
if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
  httpsConfig = {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  };
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, __dirname, '');
  const backendTarget = env.VITE_BACKEND_TARGET || 'http://localhost:4001';

  return {
    plugins: [react(), tailwindcss()],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, './src'),
      },
    },
    build: {
      outDir: 'dist',
      rollupOptions: {
        input: {
          taskpane: './src/taskpane/taskpane.html',
          authDialog: './src/auth/auth-dialog.html',
          authComplete: './src/auth/auth-complete.html',
        },
      },
    },
    server: {
      port: 3000,
      strictPort: true,
      https: httpsConfig,
      cors: {
        origin: 'https://localhost:3000',
        credentials: true,
      },
      proxy: {
        // Keep /api/auth prefix — Better Auth mounts at /api/auth on the Nest server.
        '/api/auth': {
          target: backendTarget,
          changeOrigin: true,
          secure: false,
          configure: (proxy) => {
            proxy.on('proxyReq', (proxyReq) => {
              proxyReq.setHeader('x-forwarded-proto', 'https');
              proxyReq.setHeader('x-forwarded-host', 'localhost:3000');
            });
          },
          headers: backendTarget.includes('.ngrok-free.app')
            ? {
                'ngrok-skip-browser-warning': 'true',
              }
            : undefined,
        },
        '/api': {
          target: backendTarget,
          changeOrigin: true,
          secure: false,
          rewrite: (proxyPath) => proxyPath.replace(/^\/api/, ''),
          configure: (proxy) => {
            proxy.on('proxyReq', (proxyReq) => {
              proxyReq.setHeader('x-forwarded-proto', 'https');
              proxyReq.setHeader('x-forwarded-host', 'localhost:3000');
            });
          },
          headers: backendTarget.includes('.ngrok-free.app')
            ? {
                'ngrok-skip-browser-warning': 'true',
              }
            : undefined,
        },
      },
    },
  };
});
