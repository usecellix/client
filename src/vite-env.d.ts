/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_API_BASE_URL?: string;
  readonly VITE_BACKEND_TARGET?: string;
  readonly VITE_MARKETING_SITE_URL?: string;
  readonly VITE_UPGRADE_URL?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
