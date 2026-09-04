import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // The Docs/Sheets editors embed Univer, an imperative canvas library that
  // mounts its own React root into a container div and tears down shared
  // global state on dispose(). StrictMode's dev-only double-invoke
  // (mount -> cleanup -> mount) fights that: deferring the dispose call to
  // dodge the "synchronous unmount while rendering" warning corrupts the
  // second mount instead (verified empirically), so dispose has to stay
  // synchronous — which is exactly what triggers the warning. Disabling
  // StrictMode removes the double-invoke (and the warning) entirely; it
  // doesn't change production behavior, which never double-invokes anyway.
  reactStrictMode: false,
};

export default nextConfig;
