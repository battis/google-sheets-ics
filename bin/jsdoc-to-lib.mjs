import ts from 'typescript';
import fs from 'fs';
import path from 'path';
import * as url from 'url';
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

// @see https://stackoverflow.com/a/61019690
export function extractWithComment(fileNames, options = {}) {
  const program = ts.createProgram(fileNames, options);
  const comments = [];
  program.getSourceFiles().forEach((source) => {
    if (!source.isDeclarationFile) {
      ts.forEachChild(source, (node) => {
        const commentRanges = ts.getLeadingCommentRanges(
          source.getFullText(),
          node.getFullStart()
        );
        if (commentRanges && commentRanges.length) {
          comments.push({
            name: commentRanges.map(
              (r) =>
                source
                  .getFullText()
                  .slice(r.end + 1)
                  .match(/function (.+)\(/)[1]
            ),
            jsdoc: commentRanges
              .map((r) => source.getFullText().slice(r.pos, r.end))
              .filter((c) => /^\/\*\*/.test(c))
          });
        }
      });
    }
  });
  return comments.filter((c) => /@customfunction/.test(c.jsdoc));
}

const comments = extractWithComment([
  path.resolve(__dirname, '../src/CustomFunctions.global.ts')
]);

let build = fs
  .readFileSync(path.join(__dirname, '../build/main-bundle.js'))
  .toString();
comments.forEach((comment) => {
  const paramDocs = comment.jsdoc
    .toString()
    .match(/\* @param \{[^}]+\}\s+\w+(\s|=)/g);
  const params = paramDocs.map((p) => p.match(/\}\s+(\w+)(\s|=)/)[1]).join(',');
  build = build.replace(
    new RegExp(`(function ${comment.name})\\(\\)\\{\\}`),
    `\n${comment.jsdoc}\n$1(${params}){}`
  );
  console.log(`Documented ${comment.name}()`);
});
fs.writeFileSync(path.join(__dirname, '../build/main-bundle.js'), build);
console.log('Updated main-bundle.js');
