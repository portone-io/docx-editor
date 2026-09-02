const changelogGithub = require("@changesets/changelog-github");

const github = changelogGithub.default ?? changelogGithub;

/**
 * The generator thanks whoever opened the pull request, which reads oddly for the maintainer
 * writing their own release notes. An entry names its author instead, the way GitHub's own
 * generated release notes do.
 *
 * A line the pattern does not match is left as it was written, so a change to the generator's
 * format cannot fail a release over its attribution.
 */
const CREDIT = /Thanks (\[@.+?\]\([^)]+\))! - /;

function attributed(line) {
  return line.replace(CREDIT, "By $1 - ");
}

module.exports = {
  getDependencyReleaseLine: github.getDependencyReleaseLine,
  getReleaseLine: async (...args) =>
    attributed(await github.getReleaseLine(...args)),
  attributed,
};
