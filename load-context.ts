import { type PlatformProxy } from 'wrangler';

type Cloudflare = Omit<PlatformProxy<Env>, 'dispose'>;

declare module '@remix-run/cloudflare' {
  interface AppLoadContext {
    cloudflare: Cloudflare;
  }
}

// For Vercel compatibility: provide a function to create compatible context
export function createLoadContext(env?: Record<string, any>) {
  // If we're on Vercel, provide a compatible context
  if (process.env.VERCEL) {
    return {
      cloudflare: {
        env: env || process.env as Record<string, any>,
        ctx: undefined,
        cf: undefined,
        caches: undefined,
      },
    };
  }
  // Otherwise, return undefined to use default Cloudflare context
  return undefined;
}
