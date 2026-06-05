import fs from "node:fs/promises";
import vm from "node:vm";

const babelCode = await fs.readFile("public/vendor/babel.min.js", "utf8");
const context = {};
vm.createContext(context);
vm.runInContext(babelCode, context);

const source = await fs.readFile("src/app.tsx", "utf8");
const compiled = context.Babel.transform(source, {
  filename: "app.tsx",
  presets: ["typescript", "react"],
  comments: true,
}).code;

const indexHtml = `<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Dashboard - Vendas E-commerce</title>
    <link rel="stylesheet" href="./styles.css" />
  </head>
  <body>
    <div id="root"></div>

    <script src="./vendor/react.production.min.js"></script>
    <script src="./vendor/react-dom.production.min.js"></script>
    <script src="./vendor/chart.umd.min.js"></script>
    <script src="./vendor/xlsx.full.min.js"></script>
    <script src="./base-dashboard-data.js"></script>
    <script src="./app.compiled.js"></script>
  </body>
</html>
`;

await fs.writeFile(
  "public/app.compiled.js",
  `${compiled}\n//# sourceURL=src/app.tsx\n`,
  "utf8",
);
await fs.copyFile("src/styles.css", "public/styles.css");
await fs.writeFile("public/index.html", indexHtml, "utf8");
