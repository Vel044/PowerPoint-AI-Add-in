export interface ProviderEnv {
  ANTHROPIC_AUTH_TOKEN: string;
  ANTHROPIC_BASE_URL: string;
  ANTHROPIC_DEFAULT_HAIKU_MODEL?: string;
  ANTHROPIC_DEFAULT_OPUS_MODEL?: string;
  ANTHROPIC_DEFAULT_SONNET_MODEL?: string;
  ANTHROPIC_MODEL?: string;
  API_TIMEOUT_MS?: string;
  CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC?: string;
}

export interface Provider {
  name: string;
  env: ProviderEnv;
}

export interface ProvidersConfig {
  activeProvider: string;
  providers: Record<string, Provider>;
}

const STORAGE_KEY = "claude-for-office.providers";
const DEFAULT_URL = "/config/providers.json";

let cached: ProvidersConfig | null = null;

export async function loadConfig(url = DEFAULT_URL): Promise<ProvidersConfig> {
  if (cached) return cached;

  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) {
    try {
      cached = JSON.parse(stored) as ProvidersConfig;
      return cached;
    } catch {
      localStorage.removeItem(STORAGE_KEY);
    }
  }

  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) {
    throw new Error(
      `无法加载配置文件 ${url} (HTTP ${res.status})。请将 config/providers.example.json 复制为 providers.json 并填入 API Key。`
    );
  }
  cached = (await res.json()) as ProvidersConfig;
  return cached;
}

export function saveConfig(config: ProvidersConfig): void {
  cached = config;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
}

export function getActiveProvider(config: ProvidersConfig): Provider {
  const p = config.providers[config.activeProvider];
  if (!p) throw new Error(`activeProvider "${config.activeProvider}" 不存在`);
  return p;
}

export type ModelTier = "haiku" | "sonnet" | "opus";

export function resolveModel(provider: Provider, tier: ModelTier): string {
  const env = provider.env;
  if (env.ANTHROPIC_MODEL) return env.ANTHROPIC_MODEL;
  if (tier === "haiku" && env.ANTHROPIC_DEFAULT_HAIKU_MODEL) return env.ANTHROPIC_DEFAULT_HAIKU_MODEL;
  if (tier === "opus" && env.ANTHROPIC_DEFAULT_OPUS_MODEL) return env.ANTHROPIC_DEFAULT_OPUS_MODEL;
  if (env.ANTHROPIC_DEFAULT_SONNET_MODEL) return env.ANTHROPIC_DEFAULT_SONNET_MODEL;
  return "claude-sonnet-4-6";
}

export function getTimeoutMs(provider: Provider): number {
  const v = provider.env.API_TIMEOUT_MS;
  const n = v ? parseInt(v, 10) : NaN;
  return Number.isFinite(n) ? n : 120000;
}

export function clearCache(): void {
  cached = null;
}
