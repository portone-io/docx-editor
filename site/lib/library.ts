import manifest from "../../package.json";
import siteManifest from "../package.json";

export const libraryName = manifest.name;
export const libraryDescription = manifest.description;
// The badge names the version the demo below it runs, so it reads the pin rather than the
// version the repository is working towards. `pnpm check:demo-library` holds the two together.
export const libraryVersion = siteManifest.dependencies["@portone/docx-editor"];
export const repositoryUrl = "https://github.com/portone-io/docx-editor";
export const npmUrl = `https://www.npmjs.com/package/${libraryName}`;
export const siteUrl = manifest.homepage;
export const googleSiteVerification =
  "2hn5FQAHbj-6fwc6W85KxkCtnx2CbPPDAqzfq38TIxQ";
export const docsRoute = "/docs";
export const openGraphImage = {
  url: "/og.png",
  width: 1200,
  height: 630,
  alt: "A Word document open in the docx-editor live demo",
};
