const fs = require("fs");
const path = require("path");
const archiver = require("archiver");
const JSZip = require("jszip");
const DOCX = require("../utils/docx-params");

const unzipDocx = async (docxPath, outputDir) => {
  try {
    const data = await fs.promises.readFile(docxPath);
    const zip = await JSZip.loadAsync(data);

    await Promise.all(
      Object.keys(zip.files).map(async (relativePath) => {
        const file = zip.files[relativePath];
        const fullPath = path.join(outputDir, relativePath);

        if (file.dir) {
          await fs.promises.mkdir(fullPath, { recursive: true });
        } else {
          const content = await file.async("nodebuffer");
          await fs.promises.mkdir(path.dirname(fullPath), { recursive: true });
          await fs.promises.writeFile(fullPath, content);
        }
      })
    );

    console.log(`${docxPath} descompactado em ${outputDir}`);
  } catch (err) {
    throw new Error(`Erro ao descompactar ${docxPath}: ${err.message}`);
  }
};

const mergeDocxFolders = async (files, outputFilePath) => {
  if (!Array.isArray(files) || files.length === 0) {
    throw new Error(DOCX.ERROR_INVALID_FOLDER_LIST);
  }

  const tempFolders = [];
  for (const file of files) {
    const tempDir = path.join(
      DOCX.TEMP_DIR_MERGE,
      path.basename(file, ".docx") + "_" + Date.now()
    );
    await unzipDocx(file, tempDir);
    tempFolders.push(tempDir);
  }

  const TEMP_DIR = path.join(DOCX.TEMP_DIR_MERGE, "merged_" + Date.now());
  if (!fs.existsSync(TEMP_DIR)) {
    fs.mkdirSync(TEMP_DIR, { recursive: true });
  }

  const mediaIdMappings = {};
  let nextImageId = 1;

  copyFolderRecursiveSync(tempFolders[0], TEMP_DIR);
  
  for (let i = 1; i < tempFolders.length; i++) {
    mediaIdMappings[i] = {};
    await mergeDocxContent(tempFolders[i], TEMP_DIR, mediaIdMappings[i], nextImageId);
    
    const sourceMediaFolder = path.join(tempFolders[i], DOCX.MEDIA_FOLDER);
    if (fs.existsSync(sourceMediaFolder)) {
      const mediaFiles = fs.readdirSync(sourceMediaFolder);
      nextImageId += mediaFiles.length;
    }
  }

  const output = fs.createWriteStream(outputFilePath);
  const archive = archiver("zip", { zlib: { level: DOCX.COMPRESSION_LEVEL } });

  return new Promise((resolve, reject) => {
    output.on("close", () => {
      console.log(
        DOCX.SUCCESS_MERGE.replace("{outputFilePath}", outputFilePath)
      );
      fs.rmSync(DOCX.TEMP_DIR_MERGE, { recursive: true, force: true });
      resolve();
    });

    archive.on("error", (err) => reject(err));
    archive.pipe(output);
    archive.directory(TEMP_DIR, false);
    archive.finalize();
  });
};

const copyFolderRecursiveSync = (source, target) => {
  if (!fs.existsSync(target)) {
    fs.mkdirSync(target, { recursive: true });
  }

  fs.readdirSync(source).forEach((file) => {
    const srcPath = path.join(source, file);
    const destPath = path.join(target, file);

    if (fs.lstatSync(srcPath).isDirectory()) {
      copyFolderRecursiveSync(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  });
};

const mergeMedia = (sourceFolder, targetFolder, mediaMapping, startImageId) => {
  const sourceMedia = path.join(sourceFolder, DOCX.MEDIA_FOLDER);
  const targetMedia = path.join(targetFolder, DOCX.MEDIA_FOLDER);

  if (!fs.existsSync(sourceMedia)) return {};

  if (!fs.existsSync(targetMedia)) {
    fs.mkdirSync(targetMedia, { recursive: true });
  }

  let currentImageId = startImageId;
  
  fs.readdirSync(sourceMedia).forEach((file) => {
    const srcFile = path.join(sourceMedia, file);
    
    const fileExt = path.extname(file);
    const oldFileName = path.basename(file);
    const newFileName = `image${currentImageId}_${Date.now()}${fileExt}`;
    const destFile = path.join(targetMedia, newFileName);
    
    mediaMapping[oldFileName] = newFileName;
    
    fs.copyFileSync(srcFile, destFile);
    currentImageId++;
  });
  
  return mediaMapping;
};

const updateRelationships = (sourceFolder, targetFolder, mediaMapping) => {
  const documentRelsPath = "word/_rels/document.xml.rels";
  
  const sourceDocRelsPath = path.join(sourceFolder, documentRelsPath);
  const targetDocRelsPath = path.join(targetFolder, documentRelsPath);

  if (!fs.existsSync(sourceDocRelsPath)) return;
  
  if (!fs.existsSync(targetDocRelsPath)) {
    fs.mkdirSync(path.dirname(targetDocRelsPath), { recursive: true });
    fs.writeFileSync(targetDocRelsPath, 
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n` +
      `</Relationships>`, 
      DOCX.ENCODING);
  }

  let sourceRels = fs.readFileSync(sourceDocRelsPath, DOCX.ENCODING);
  let targetRels = fs.readFileSync(targetDocRelsPath, DOCX.ENCODING);
  
  const existingRIds = [...targetRels.matchAll(/Id="(rId\d+)"/g)].map(m => parseInt(m[1].replace('rId', '')));
  let maxRId = existingRIds.length > 0 ? Math.max(...existingRIds) : 0;

  const sourceRelMatches = [...sourceRels.matchAll(/<Relationship [^>]*>/g)];
  
  sourceRelMatches.forEach((match) => {
    const rel = match[0];
    
    if (rel.includes('media/')) {
      const idMatch = rel.match(/Id="([^"]+)"/);
      const targetMatch = rel.match(/Target="([^"]+)"/);
      
      if (idMatch && targetMatch) {
        const oldId = idMatch[1];
        const oldTarget = targetMatch[1];
        const oldFileName = path.basename(oldTarget);
        
        if (mediaMapping[oldFileName]) {
          maxRId++;
          const newId = `rId${maxRId}`;
          const newTarget = oldTarget.replace(oldFileName, mediaMapping[oldFileName]);

          const newRel = rel
            .replace(/Id="([^"]+)"/, `Id="${newId}"`)
            .replace(/Target="([^"]+)"/, `Target="${newTarget}"`);
          
          mediaMapping[oldId] = newId;
          
          targetRels = targetRels.replace(
            /<\/Relationships>/,
            `${newRel}\n</Relationships>`
          );
        }
      }
    } else {
      if (!targetRels.includes(rel)) {
        targetRels = targetRels.replace(
          /<\/Relationships>/,
          `${rel}\n</Relationships>`
        );
      }
    }
  });

  fs.writeFileSync(targetDocRelsPath, targetRels, DOCX.ENCODING);
  
  const mainRelsPath = "_rels/.rels";
  const sourceRelsPath = path.join(sourceFolder, mainRelsPath);
  const targetRelsPath = path.join(targetFolder, mainRelsPath);

  if (fs.existsSync(sourceRelsPath) && fs.existsSync(targetRelsPath)) {
    let sourceMainRels = fs.readFileSync(sourceRelsPath, DOCX.ENCODING);
    let targetMainRels = fs.readFileSync(targetRelsPath, DOCX.ENCODING);

    const relationshipRegex = /<Relationship [^>]*>/g;
    const relationships = sourceMainRels.match(relationshipRegex) || [];
    
    relationships.forEach((rel) => {
      if (!targetMainRels.includes(rel)) {
        targetMainRels = targetMainRels.replace(
          /<\/Relationships>/,
          `${rel}\n</Relationships>`
        );
      }
    });

    fs.writeFileSync(targetRelsPath, targetMainRels, DOCX.ENCODING);
  }
};

const mergeStylesAndSettings = (sourceFolder, targetFolder) => {
  const styleFiles = [
    "word/styles.xml",
    "word/numbering.xml",
    "word/settings.xml",
    "word/fontTable.xml",
    "word/webSettings.xml",
    "word/theme/theme1.xml"
  ];

  styleFiles.forEach((file) => {
    const sourcePath = path.join(sourceFolder, file);
    const targetPath = path.join(targetFolder, file);

    if (fs.existsSync(sourcePath)) {
      fs.mkdirSync(path.dirname(targetPath), { recursive: true });
      
      if (!fs.existsSync(targetPath)) {
        fs.copyFileSync(sourcePath, targetPath);
      }
    }
  });
};

const updateDocumentXmlImageRefs = (xmlContent, mediaMapping) => {
  let updatedXml = xmlContent;
  
  for (const oldRelId in mediaMapping) {
    const newRelId = mediaMapping[oldRelId];
    
    const patterns = [
      new RegExp(`r:id="${oldRelId}"`, 'g'),
      new RegExp(`r:id='${oldRelId}'`, 'g'),
      new RegExp(`r:embed="${oldRelId}"`, 'g'),
      new RegExp(`r:embed='${oldRelId}'`, 'g'),
      new RegExp(`a:blip[^>]*r:embed="${oldRelId}"`, 'g'),
      new RegExp(`a:blip[^>]*r:embed='${oldRelId}'`, 'g'),
      new RegExp(`relationships:id="${oldRelId}"`, 'g'),
      new RegExp(`relationships:id='${oldRelId}'`, 'g')
    ];
    
    patterns.forEach(pattern => {
      updatedXml = updatedXml.replace(pattern, (match) => {
        return match.replace(oldRelId, newRelId);
      });
    });
  }
  
  return updatedXml;
};
const mergeDocxContent = async (sourceFolder, targetFolder, mediaMapping, startImageId) => {
  const documentXmlPath = "word/document.xml";
  const docXmlPath1 = path.join(targetFolder, documentXmlPath);
  const docXmlPath2 = path.join(sourceFolder, documentXmlPath);

  if (fs.existsSync(docXmlPath1) && fs.existsSync(docXmlPath2)) {
    const content1 = fs.readFileSync(docXmlPath1, DOCX.ENCODING);
    const content2 = fs.readFileSync(docXmlPath2, DOCX.ENCODING);

    mergeMedia(sourceFolder, targetFolder, mediaMapping, startImageId);
    
    updateRelationships(sourceFolder, targetFolder, mediaMapping);
    
    const updatedContent2 = updateDocumentXmlImageRefs(content2, mediaMapping);
    
    const extractBody = (content) => {
      const bodyMatch = content.match(/<w:body>([\s\S]*?)<\/w:body>/);
      if (!bodyMatch) return { body: "", sectPr: "" };
      
      const bodyContent = bodyMatch[1];
      const sectPrMatch = bodyContent.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
      
      if (sectPrMatch) {
        return {
          body: bodyContent.replace(sectPrMatch[0], ""),
          sectPr: sectPrMatch[0]
        };
      }
      
      return { body: bodyContent, sectPr: "" };
    };
    
    const doc1Parts = extractBody(content1);
    const doc2Parts = extractBody(updatedContent2);
    
    const sectPr = doc1Parts.sectPr || doc2Parts.sectPr;
    
    const mergedBody = doc1Parts.body + doc2Parts.body + sectPr;
    
    const mergedContent = content1.replace(
      /<w:body>[\s\S]*?<\/w:body>/,
      `<w:body>${mergedBody}</w:body>`
    );

    fs.writeFileSync(docXmlPath1, mergedContent, DOCX.ENCODING);
    
    mergeStylesAndSettings(sourceFolder, targetFolder);
  }
};

module.exports = { mergeDocxFolders };