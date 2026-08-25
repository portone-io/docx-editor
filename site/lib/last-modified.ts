import { execFileSync } from "node:child_process";

const commitTime = (path: string) => {
  const stdout = execFileSync(
    "git",
    ["log", "-1", "--format=%cI", "--", path],
    {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "ignore"],
    }
  ).trim();
  return stdout ? new Date(stdout) : undefined;
};

// Deploys clone shallowly, so a file untouched within the fetched history has no
// commit to report, and git itself may be absent outside a checkout. Leaving
// lastmod off a URL tells crawlers less than a wrong date tells them wrongly, so
// every failure resolves to undefined rather than to the build time.
export const lastModified = (path: string | undefined) => {
  if (!path) return undefined;
  try {
    return commitTime(path);
  } catch {
    return undefined;
  }
};
