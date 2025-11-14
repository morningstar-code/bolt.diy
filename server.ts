import { createLoadContext } from './load-context';

/**
 * This function is called by Vercel's Remix integration to create the load context
 * for each request. It provides a Cloudflare-compatible context structure
 * by mapping process.env to cloudflare.env.
 */
export function getLoadContext(event: any) {
  // Create a compatible load context for Vercel
  // Pass process.env to map to cloudflare.env structure
  return createLoadContext(process.env);
}

