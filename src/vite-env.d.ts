/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_EDS_API_URL: string;
  readonly VITE_EDS_APP_TOKEN: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
