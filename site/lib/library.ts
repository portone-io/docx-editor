import manifest from "../../package.json";

export const libraryName = manifest.name;
export const libraryDescription = manifest.description;
export const libraryVersion = manifest.version;
export const repositoryUrl = "https://github.com/portone-io/docx-editor";
export const siteUrl = manifest.homepage;
// Search Console rechecks ownership over time, so this has to keep being
// served rather than being dropped once the property turns green.
export const googleSiteVerification =
  "2hn5FQAHbj-6fwc6W85KxkCtnx2CbPPDAqzfq38TIxQ";
export const docsRoute = "/docs";
