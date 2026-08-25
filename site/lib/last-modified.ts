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

export const lastModified = (path: string | undefined) => {
  if (!path) return undefined;
  try {
    return commitTime(path);
  } catch {
    return undefined;
  }
};
