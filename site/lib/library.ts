import manifest from "../../package.json";
import installed from "../node_modules/@portone/docx-editor/package.json";

export const libraryName = manifest.name;
export const libraryDescription = manifest.description;
// The badge names the version the demo below it runs, so it reads the installed library
// rather than the version the repository is working towards: the sources on main, the
// published release on the production branch. `pnpm check:demo-library` holds the
// manifests and the installation together.
export const libraryVersion = installed.version;
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
