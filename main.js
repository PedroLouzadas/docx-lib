const fs = require("fs");
const path = require("path");
const { execSync, spawn } = require("child_process");

const docsToMerge = [
  "./doc1.docx",
  "./doc2.docx",
  "./doc3.docx",
  "./doc4.docx",
  "./doc5.docx",
  "./doc6.docx",
  "./doc7.docx",
  "./doc8.docx",
  "./doc9.docx",
  "./doc10.docx"
];

const outputPath = "./merged.pdf";

const missingFiles = docsToMerge.filter((file) => !fs.existsSync(file));
if (missingFiles.length > 0) {
  console.error("Arquivos não encontrados:", missingFiles.join(", "));
  process.exit(1);
}
const sofficePath = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

function convertDocxToPdf(inputPath) {
  try {
    console.log(`Convertendo ${inputPath} para PDF...`);

    const cmd = `"${sofficePath}" --headless --convert-to pdf "${inputPath}" --outdir "./saida"`;
    execSync(cmd);
  } catch (error) {
    console.error(`Erro ao converter ${inputPath}:`, error.message);
  }
}

docsToMerge.forEach(convertDocxToPdf);

const pdfsToMerge = docsToMerge.map((file) => {
  const filename = path.basename(file, ".docx") + ".pdf";
  return path.join("saida", filename);
});

function runPythonMerge() {
  return new Promise((resolve, reject) => {
    const args = ["merge_pdfs.py", outputPath, ...pdfsToMerge];
    const python = spawn("py", args);

    python.stdout.on("data", (data) => console.log(data.toString()));
    python.stderr.on("data", (data) => console.error(data.toString()));

    python.on("close", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`Erro ao unir PDFs. Código de saída: ${code}`));
    });
  });
}

runPythonMerge()
  .then(() => {
    console.log("PDFs mesclados com sucesso!");
    console.log(`Arquivo final: ${path.resolve(outputPath)}`);
  })
  .catch((error) => {
    console.error("Erro na mesclagem:", error.message);
    process.exit(1);
  });
