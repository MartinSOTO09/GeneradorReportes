async function generarZIP(data) {
    const zip = new JSZip();

    try {
        // Seleccionar generadores según el tipo de DB (soporta wrappers por tipo)
        const db = (data && data.db_type) ? String(data.db_type).toLowerCase() : 'oracle';
        const rpFnName = db === 'postgres' ? 'generarReleasePlan_postgres' : (db === 'linux' ? 'generarReleasePlan_linux' : 'generarReleasePlan');
        const teFnName = db === 'postgres' ? 'generarTestEvidence_postgres' : (db === 'linux' ? 'generarTestEvidence_linux' : 'generarTestEvidence');

        const rpFn = (typeof window !== 'undefined' && typeof window[rpFnName] === 'function') ? window[rpFnName] : (typeof globalThis !== 'undefined' && typeof globalThis[rpFnName] === 'function' ? globalThis[rpFnName] : generarReleasePlan);
        const teFn = (typeof window !== 'undefined' && typeof window[teFnName] === 'function') ? window[teFnName] : (typeof globalThis !== 'undefined' && typeof globalThis[teFnName] === 'function' ? globalThis[teFnName] : generarTestEvidence);

        // Generar documentos Word en paralelo usando las funciones seleccionadas
        const [releasePlanDoc, testEvidenceDoc] = await Promise.all([
            rpFn(data),
            teFn(data)
        ]);

        // Asegurar que los archivos agregados a JSZip sean ArrayBuffer o Uint8Array (compatibilidad)
        const addDocToZip = async (name, doc) => {
            try {
                if (doc instanceof Blob) {
                    const buf = await doc.arrayBuffer();
                    zip.file(name, buf);
                    console.debug(`Agregado a ZIP (Blob -> ArrayBuffer): ${name} (${buf.byteLength} bytes)`);
                } else if (doc && typeof doc === 'object' && doc.arrayBuffer) {
                    const buf = await doc.arrayBuffer();
                    zip.file(name, buf);
                    console.debug(`Agregado a ZIP (arrayBuffer()): ${name} (${buf.byteLength} bytes)`);
                } else {
                    // Si es un ArrayBuffer/Buffer o string
                    zip.file(name, doc);
                    console.debug(`Agregado a ZIP (raw): ${name}`);
                }
            } catch (e) {
                console.error(`Error agregando ${name} al ZIP:`, e);
                throw e;
            }
        };

        await addDocToZip(`RP - ${data.solman}.docx`, releasePlanDoc);
        await addDocToZip(`TE - ${data.solman}.docx`, testEvidenceDoc);
    } catch (error) {
        console.error("Fallo al generar los documentos .docx:", error);
        // Continuar permitiendo adjuntar respaldos aunque fallen los docx
    }

    // Adjuntar respaldos directamente en la raíz del ZIP
    try {
        if (data.respaldos && data.respaldos.length > 0) {
            const files = Array.from(data.respaldos);
            for (let file of files) {
                try {
                    const content = await file.arrayBuffer();
                    zip.file(file.name, content);
                } catch (e) {
                    console.error(`Error leyendo archivo de respaldo ${file.name}:`, e);
                }
            }
        }
    } catch (error) {
        console.error("Fallo al adjuntar respaldos:", error);
    }

    // Generar ZIP con nombre del Solman
    try {
        const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
        // Guardar exactamente con el nombre solicitado (sin fallbacks)
        const filename = `${data.solman}.zip`;
        try {
            console.log(`Requesting download with filename: "${filename}", blob size: ${blob.size} bytes`);
        } catch (e) {
            // ignore logging errors
        }

        // Fallback download using anchor to try to avoid external renaming by handlers
        try {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            // Required for Firefox in some cases
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 2000);
        } catch (e) {
            // If anchor approach fails, fallback to FileSaver (if available)
            if (typeof saveAs === 'function') saveAs(blob, filename);
        }
    } catch (error) {
        console.error("Fallo al generar el ZIP:", error);
    }
}
