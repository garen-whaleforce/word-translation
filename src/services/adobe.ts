/**
 * Adobe PDF Services - PDF to DOCX conversion
 */

import {
  ServicePrincipalCredentials,
  PDFServices,
  MimeType,
  ExportPDFJob,
  ExportPDFParams,
  ExportPDFTargetFormat,
  ExportPDFResult,
} from "@adobe/pdfservices-node-sdk";
import * as fs from "fs";
import * as path from "path";

/**
 * Convert a PDF file to DOCX using Adobe PDF Services API
 * @param inputPath - Path to the input PDF file
 * @param outputPath - Path where the output DOCX will be saved
 */
export async function convertPdfToDocx(
  inputPath: string,
  outputPath: string
): Promise<void> {
  const clientId = process.env.PDF_SERVICES_CLIENT_ID;
  const clientSecret = process.env.PDF_SERVICES_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    throw new Error(
      "PDF_SERVICES_CLIENT_ID and PDF_SERVICES_CLIENT_SECRET must be set"
    );
  }

  // Create credentials
  const credentials = new ServicePrincipalCredentials({
    clientId,
    clientSecret,
  });

  // Create PDF Services instance
  const pdfServices = new PDFServices({ credentials });

  // Read the PDF file
  const inputStream = fs.createReadStream(inputPath);

  // Upload the PDF asset
  const inputAsset = await pdfServices.upload({
    readStream: inputStream,
    mimeType: MimeType.PDF,
  });

  // Set export parameters for DOCX
  const params = new ExportPDFParams({
    targetFormat: ExportPDFTargetFormat.DOCX,
  });

  // Create and submit the export job
  const job = new ExportPDFJob({ inputAsset, params });
  const pollingURL = await pdfServices.submit({ job });

  // Poll for job completion
  const pdfServicesResponse = await pdfServices.getJobResult({
    pollingURL,
    resultType: ExportPDFResult,
  });

  // Get the result asset
  const resultAsset = pdfServicesResponse.result?.asset;
  if (!resultAsset) {
    throw new Error("No result asset returned from Adobe PDF Services");
  }

  // Download the result
  const streamAsset = await pdfServices.getContent({ asset: resultAsset });
  const readStream = streamAsset.readStream;

  // Ensure output directory exists
  const outputDir = path.dirname(outputPath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Write to output file
  const writeStream = fs.createWriteStream(outputPath);

  return new Promise((resolve, reject) => {
    readStream.pipe(writeStream);
    writeStream.on("finish", () => {
      console.log(`PDF converted to DOCX: ${outputPath}`);
      resolve();
    });
    writeStream.on("error", reject);
    readStream.on("error", reject);
  });
}
