import type { MetadataRoute } from "next";

const siteUrl = "https://dockaro.com";

export default function sitemap(): MetadataRoute.Sitemap {
  const routes = ["", "/pricing", "/api-docs", "/legal/privacy", "/legal/terms", "/status"];

  return routes.map((route) => ({
    url: `${siteUrl}${route}`,
    lastModified: new Date(),
    changeFrequency: route === "" ? "weekly" : "monthly",
    priority: route === "" ? 1 : 0.6,
  }));
}
