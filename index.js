import express from 'express';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { PDFDocument } from 'pdf-lib';
import QRCode from 'qrcode';
import { pdf } from "pdf-to-img";
import { Jimp } from 'jimp';
import jsQR from 'jsqr';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { promisify } from 'util';
import { exec } from 'child_process';
import os from 'os';

const execPromise = promisify(exec);
const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs');

const app = express();
const upload = multer({ dest: os.tmpdir() }); 

// 1. CONVERSIÓN MEJORADA (Protección contra bloqueos y perfiles)
const convertToPdf = async (inputPath, outputDir) => {
    const isWindows = process.platform === "win32";
    // Agregamos un entorno de instalación temporal para evitar conflictos si LibreOffice está abierto
    const command = isWindows 
        ? `"C:\\Program Files\\LibreOffice\\program\\soffice.exe" --headless "-env:UserInstallation=file:///${os.tmpdir().replace(/\\/g, '/')}/libre_profile_${Date.now()}" --convert-to pdf "${inputPath}" --outdir "${outputDir}"`
        : `libreoffice --headless --convert-to pdf "${inputPath}" --outdir "${outputDir}"`;

    try {
        await execPromise(command);
        const nombreArchivo = path.parse(inputPath).name + ".pdf";
        const finalPath = path.resolve(outputDir, nombreArchivo);
        if (!fs.existsSync(finalPath)) throw new Error("El PDF no se generó.");
        return finalPath;
    } catch (err) {
        throw new Error(`Fallo en LibreOffice: ${err.message}`);
    }
};

app.post('/generar-titulo', upload.fields([
    { name: 'titulo_pdf', maxCount: 1 },
    { name: 'plantilla_docx', maxCount: 1 }
]), async (req, res) => {
    
    if (!req.files['titulo_pdf'] || !req.files['plantilla_docx']) {
        return res.status(400).send('Archivos faltantes.');
    }

    const pdfEntradaPath = req.files['titulo_pdf'][0].path;
    const plantillaPath = req.files['plantilla_docx'][0].path;
    const outputDir = os.tmpdir();
    
    const tempDocxPath = path.resolve(outputDir, `Doc_${Date.now()}.docx`);
    const imagenTempQR = path.resolve(outputDir, `qr_${Date.now()}.png`);

    try {
        const data = new Uint8Array(fs.readFileSync(pdfEntradaPath));
        const loadingTask = pdfjsLib.getDocument({ data, disableWorker: true, verbosity: 0 });
        const pdfDoc = await loadingTask.promise;
        
        let rawLines = [];
        for (let i = 1; i <= pdfDoc.numPages; i++) {
            const page = await pdfDoc.getPage(i);
            const textContent = await page.getTextContent();
            const items = textContent.items.sort((a, b) => b.transform[5] - a.transform[5]);

            let currentLine = "";
            let lastY = -1;
            for (const item of items) {
                if (Math.abs(lastY - item.transform[5]) > 2 && lastY !== -1) {
                    rawLines.push(currentLine.trim());
                    currentLine = item.str;
                } else {
                    currentLine += (currentLine === "" ? "" : "  ") + item.str;
                }
                lastY = item.transform[5];
            }
            rawLines.push(currentLine.trim());
        }

        // 2. EXTRACCIÓN ROBUSTA DE LÍNEAS (Fecha completa)
        let lines = [];
        for (let j = 0; j < rawLines.length; j++) {
            let line = rawLines[j].trim();
            if (!line) continue;

            if (line.startsWith("Sello digital")) {
                let fullSello = line;
                while (j + 1 < rawLines.length && !rawLines[j + 1].includes(":") && !rawLines[j + 1].startsWith("Fecha y hora")) {
                    fullSello += rawLines[j + 1].replace(/\s+/g, '');
                    j++;
                }
                lines.push(fullSello);
            } 
            else if (line.startsWith("Fecha y hora de sellado")) {
                let fullFecha = line;
                // BUCLE MEJORADO: Continúa uniendo líneas hasta que encuentre el patrón de hora XX:XX:XX
                while (j + 1 < rawLines.length && !/\d{2}:\d{2}:\d{2}/.test(fullFecha)) {
                    let nextLine = rawLines[j + 1].trim();
                    
                    // Freno de emergencia por si el PDF no tiene la hora completa y empieza otra etiqueta
                    if (nextLine.includes("No. Certificado") || nextLine.includes("Sello digital")) break;

                    // Si el pedazo huérfano empieza con ":", lo unimos SIN espacio para reconstruir la hora
                    if (nextLine.startsWith(":")) {
                        fullFecha += nextLine;
                    } else {
                        fullFecha += " " + nextLine;
                    }
                    j++;
                }
                lines.push(fullFecha);
            }
            else { lines.push(line); }
        }

        const limpiar = (str) => str ? str.replace(/\s{2,}/g, ' ').trim() : "";

        // 3. MAPEO DE DATOS (Manejo de múltiples ":")
        let datosExtraidos = {};
        if (lines.length > 7) {
            const lCarrera = lines[3].split(/\s{2,}/);
            const lFechas = lines[4].split(/\s{2,}/);
            const lClaves = lines[6].split(/\s{2,}/);
            const lEntidad = lines[7].split(/\s{2,}/);

            datosExtraidos = {
                "Folio": lines[0],
                "CURP": lines[1],
                "Nombre del Profesionista": limpiar(lines[2]),
                "Carrera": limpiar(lCarrera[0] || ""),
                "ClaveCarrera": limpiar(lCarrera[1] || ""),
                "Fechas Inicio": lFechas[0] || "",
                "Fechas Fin": lFechas[1] || "",
                "Fechas Examen": lFechas[2] || "",
                "Institución": limpiar(lines[5]),
                "ClaveInst": lClaves[0] || "",
                "Autorización": lClaves[1] || "",
                "Entidad": lEntidad[0] || "",
                "Fecha de Expedición": lEntidad[1] || ""
            };

            lines.forEach((linea) => {
                const etiquetas = [
                    "Autoridad educativa", "No. Certificado autoridad educativa", 
                    "Sello digital autoridad educativa", "Responsable del centro educativo", 
                    "Fecha y hora de sellado", "No. Certificado del responsable del centro educativo", 
                    "Sello digital responsable"
                ];

                etiquetas.forEach(etiq => {
                    // Usamos startsWith para asegurar que agarramos la etiqueta correcta como inicio de línea
                    if (linea.startsWith(etiq)) {
                        // Extraemos todo el texto después del primer ":" 
                        const indexDosPuntos = linea.indexOf(":");
                        if (indexDosPuntos !== -1) {
                            const valor = linea.substring(indexDosPuntos + 1).trim();
                            datosExtraidos[etiq] = valor.replace(/[\r\n]+/g, " ");
                        }
                    }
                });
            });
        }
        // 4. QR Y DOCX
        const pdfImgConv = await pdf(pdfEntradaPath, { scale: 4 });
        for await (const imagen of pdfImgConv) {
            fs.writeFileSync(imagenTempQR, imagen);
            break; 
        }
        const imgParaQR = await Jimp.read(imagenTempQR);
        const qrLeido = jsQR(new Uint8ClampedArray(imgParaQR.bitmap.data), imgParaQR.bitmap.width, imgParaQR.bitmap.height);
        
        const zip = new PizZip(fs.readFileSync(plantillaPath, 'binary'));
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.render(datosExtraidos);
        fs.writeFileSync(tempDocxPath, doc.getZip().generate({ type: 'nodebuffer' }));

        const pdfGeneradoPath = await convertToPdf(tempDocxPath, outputDir);

        // 5. ESTAMPADO FINAL
        const pdfBytes = fs.readFileSync(pdfGeneradoPath);
        const pdfDocFinal = await PDFDocument.load(pdfBytes);
        const qrBuffer = await QRCode.toBuffer(qrLeido?.data || "Error", { margin: 1, errorCorrectionLevel: 'H' });
        const qrImage = await pdfDocFinal.embedPng(qrBuffer);
        
        const { height: pageHeight } = pdfDocFinal.getPages()[0].getSize();
        pdfDocFinal.getPages()[0].drawImage(qrImage, {
            x: 301 / 4,
            y: pageHeight - (2216 / 4) - ((2032 - 1616) / 4),
            width: (695 - 308) / 4,
            height: (2032 - 1616) / 4,
        });

        res.setHeader('Content-Type', 'application/pdf');
        res.send(Buffer.from(await pdfDocFinal.save()));

    } catch (error) {
        console.error("ERROR:", error.message);
        res.status(500).send('Error procesando el documento.');
    } finally {
        [pdfEntradaPath, plantillaPath, tempDocxPath, imagenTempQR].forEach(p => {
            if (p && fs.existsSync(p)) fs.unlinkSync(p);
        });
    }
});

app.listen(3005, () => console.log('Servidor en puerto 3005'));