async function generarTestEvidence(data) {
    if (typeof docx === 'undefined') {
        throw new Error("La librería docx no está cargada");
    }

    const { Document, Packer, Paragraph, TextRun, AlignmentType, Header, Table, TableRow, TableCell, WidthType, ImageRun, BorderStyle, PageBreak, TableOfContents, HeadingLevel } = docx;

    try {
        // Cargar imagen como ArrayBuffer
        const imageBuffer = await (await fetch("./img/tenaris.png")).arrayBuffer();

        // Helpers para construir partes del documento
        function makeHorizontalLine() {
            return new Paragraph({
                children: [
                    new TextRun({
                        text: "________________________________________________________________________________________________________________________________________________________",
                        size: 12,
                        font: "Calibri",
                        bold: true
                    })
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            });
        }

        function makeHeaderTable(imageBuf) {
            return new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                width: { size: 50, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ children: [new ImageRun({ data: imageBuf, transformation: { width: 130, height: 30 } })], alignment: AlignmentType.LEFT })],
                                margins: { top: 0, bottom: 0, left: 0, right: 0 },
                                borders: { top: { style: BorderStyle.NONE, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, color: "FFFFFF" }, left: { style: BorderStyle.NONE, color: "FFFFFF" }, right: { style: BorderStyle.NONE, color: "FFFFFF" } }
                            }),
                            new TableCell({
                                width: { size: 50, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ children: [new TextRun({ text: "IT Tenaris", bold: true, size: 48, font: "Calibri" })], alignment: AlignmentType.RIGHT })],
                                margins: { top: 0, bottom: 0, left: 0, right: 0 },
                                borders: { top: { style: BorderStyle.NONE, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, color: "FFFFFF" }, left: { style: BorderStyle.NONE, color: "FFFFFF" }, right: { style: BorderStyle.NONE, color: "FFFFFF" } }
                            })
                        ]
                    })
                ]
            });
        }

        function makeSecondPageHeader(info) {
            return new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: `Release Number: ${info.solman}`, size: 18, font: 'Calibri' })], alignment: AlignmentType.LEFT })], margins: { top: 100, bottom: 100, left: 100, right: 100 }, borders: { top: { style: BorderStyle.NONE, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, color: 'FFFFFF' }, left: { style: BorderStyle.NONE, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, color: 'FFFFFF' } } }),
                            new TableCell({ width: { size: 34, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: `Author: ${info.usuario.nombre}`, size: 18, font: 'Calibri' })], alignment: AlignmentType.CENTER })], margins: { top: 100, bottom: 100, left: 100, right: 100 }, borders: { top: { style: BorderStyle.NONE, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, color: 'FFFFFF' }, left: { style: BorderStyle.NONE, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, color: 'FFFFFF' } } }),
                            new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: `Date: ${new Date().toLocaleDateString('es-ES')}`, size: 18, font: 'Calibri' })], alignment: AlignmentType.RIGHT })], margins: { top: 100, bottom: 100, left: 100, right: 100 }, borders: { top: { style: BorderStyle.NONE, color: 'FFFFFF' }, bottom: { style: BorderStyle.NONE, color: 'FFFFFF' }, left: { style: BorderStyle.NONE, color: 'FFFFFF' }, right: { style: BorderStyle.NONE, color: 'FFFFFF' } } })
                        ]
                    })
                ]
            });
        }

        /* Tabla de 2 columnas */
        function makeStyledTwoColumnTable2(rowsData) {
            return new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                    bottom: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                    left: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                    right: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                    insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                    insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' }
                },
                rows: rowsData.map(([label, value], idx) => {
                    // Función para procesar el contenido y crear párrafos
                    const processContent = (content) => {
                        if (Array.isArray(content)) {
                            // Ahora las viñetas se muestran solo cuando el propio item tiene `bullet: true`.
                            // Para tener un control fino, representamos la viñeta como un TextRun pequeño '• ' precediendo el texto.
                            return content.map((item, itemIndex) => {
                                const text = (typeof item === 'object' && item !== null) ? (item.text || '') : String(item);
                                const itemBold = (typeof item === 'object' && item !== null && item.bold) ? true : (idx === 0 && itemIndex === 0);
                                const itemItalics = (typeof item === 'object' && item !== null && typeof item.italics === 'boolean') ? item.italics : true;

                                const explicitBullet = (typeof item === 'object' && item !== null && item.bullet === true);

                                const children = [];
                                if (explicitBullet) {
                                    // bullet pequeño (carácter) con tamaño reducido
                                    children.push(new TextRun({ text: '•\t', font: 'Calibri', size: 18 }));
                                }

                                children.push(new TextRun({
                                    text: text,
                                    font: 'Calibri',
                                    size: idx === 0 ? 18 : 14,
                                    bold: itemBold,
                                    color: '000000',
                                    italics: itemItalics
                                }));

                                return new Paragraph({ children, alignment: AlignmentType.LEFT });
                            });
                        } else {
                            // Si es string simple, crear un párrafo normal
                            return [new Paragraph({
                                children: [
                                    new TextRun({
                                        text: content,
                                        font: 'Calibri',
                                        size: idx === 0 ? 18 : 14,
                                        bold: idx === 0,
                                        color: '000000',
                                        italics: true
                                    })
                                ],
                                alignment: AlignmentType.LEFT
                            })];
                        }
                    };

                    return new TableRow({
                        children: [
                            // Columna izquierda
                            new TableCell({
                                width: { size: 20, type: WidthType.PERCENTAGE },
                                children: processContent(label),
                                shading: { fill: idx === 0 ? 'b6b6b6' : 'FFFFFF' },
                                margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                borders: {
                                    top: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                                    bottom: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                                    left: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                                    right: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' }
                                }
                            }),
                            // Columna derecha
                            new TableCell({
                                width: { size: 80, type: WidthType.PERCENTAGE },
                                children: processContent(value),
                                shading: { fill: idx === 0 ? 'b6b6b6' : 'FFFFFF' },
                                margins: { top: 100, bottom: 100, left: 200, right: 200 },
                                borders: {
                                    top: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                                    bottom: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                                    left: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                                    right: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' }
                                }
                            })
                        ]
                    });
                })
            });
        }

        /* Tabla dinamica con N filas y N columnas */
        function makeStyledDynamicTable(rowsData) {
            return new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: rowsData.map((row, rowIndex) =>
                    new TableRow({
                        children: row.map((cellContent) => {
                            let contentArray;
                            if (typeof cellContent === 'string') {
                                // Dividir por salto de línea
                                contentArray = cellContent.split('\n').map(line => ({ text: line }));
                            } else if (Array.isArray(cellContent)) {
                                contentArray = cellContent;
                            } else {
                                throw new Error("Celda inválida: debe ser string o array de objetos");
                            }

                            // Crear Paragraph con múltiples TextRun según el estilo
                            const paragraph = new Paragraph({
                                children: contentArray.flatMap((item, index) => {
                                    const tr = new TextRun({
                                        text: item.text,
                                        font: item.font || 'Calibri',
                                        size: item.size || 20,
                                        bold: item.bold || false,
                                        italic: item.italic || false,
                                        color: item.color || '000000',
                                        // respetar highlight si viene en los datos (e.g. 'Yellow')
                                        highlight: item.highlight ? String(item.highlight).toLowerCase() : undefined
                                    });
                                    // Agregar salto de línea si no es el último
                                    return index < contentArray.length - 1 ? [tr, new TextRun({ break: 1 })] : [tr];
                                }),
                                alignment: rowIndex === 0 ? AlignmentType.CENTER : AlignmentType.LEFT
                            });

                            return new TableCell({
                                children: [paragraph],
                                shading: rowIndex === 0 ? { fill: 'c1d59a' } : undefined,
                                margins: { top: 75, bottom: 75, left: 75, right: 75 }
                            });
                        })
                    })
                )
            });
        }


        // Crear un índice manual en formato de párrafos usando puntos generados manualmente
        // Recibe opcionalmente un array de entradas y un array de números (strings) para la columna derecha
        function makeManualTOC(entriesInput, numbersInput) {
            const entries = entriesInput || [
                '1.1    OBJECTIVE',
                '1.2    REFERENCES',
                '4.1    AFFECTED ENVIRONMENTS (CHECKMARK):',
                '4.2    AFFECTED SYSTEMS',
                '4.3    SCENARIOS WHERE CHANGES WILL BE APPLIED (CHECKMARK):',
                '4.4    HARDWARE AND SOFTWARE CONFIGURATIONS',
                '4.5    SOFTWARE INVENTORY',
                '5.1    LIST OF MAINFRAME OBJECTS',
                '5.2    LIST OF ORACLE OBJECTS',
                '5.3    LIST OF JOBS',
                '5.4    RELATED ADABAS DATA MODEL',
                '5.4.1  Other',
                '5.4.2  Cross Reference'
            ];

            const numbers = numbersInput || ['3', '3', '4', '4', '5', '5', '5', '6', '6', '7', '7', '7', '7'];

            // Longitud objetivo aproximada en caracteres para la línea completa (ajustable)
            // Aumentado para que los puntos lleguen más cerca del margen derecho
            const targetLineLength = 110;

            return entries.map((title, idx) => {
                const numberText = String(numbers[idx] || '');
                const cleanTitle = String(title).trim();
                // Construir la línea sin puntos: título + TAB + número (alineado a la derecha por tabStops)
                const leftText = cleanTitle;
                const tabPosition = 9000; // posición del tab stop derecho

                return new Paragraph({
                    children: [
                        new TextRun({ text: leftText, size: 22, font: 'Calibri' }),
                        new TextRun({ text: '\t' + numberText, size: 22, font: 'Calibri' })
                    ],
                    tabStops: [
                        { type: docx.TabStopType.RIGHT, position: tabPosition }
                    ],
                    spacing: { after: 100 }
                });
            });
        }



        // Helper: leer archivos adjuntos (.sql) y convertir a Paragraphs
        async function makeAttachmentParagraphs(files) {
            const paragraphs = [];
            if (!files || files.length === 0) return paragraphs;

            // Encabezado para la sección de respaldos
            const sectionHeading = new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun({ text: 'RESPALDOS', bold: true, size: 22, font: 'Calibri', color: '000000' })],
                spacing: { before: 200, after: 200 }
            });

            // Añadir el título al inicio de los párrafos (se agregará solo si hay archivos válidos)
            let addedSectionHeading = false;

            for (let file of files) {
                const name = file.name || 'adjunto';
                const lower = name.toLowerCase();
                // incluir solo archivos .sql
                if (!lower.endsWith('.sql')) continue;
                try {
                    // Leer contenido del archivo
                    const text = await file.text();

                    // Truncado
                    const maxChars = 1500;
                    let displayed = text;
                    let truncated = false;
                    if (text.length > maxChars) {
                        displayed = text.slice(0, maxChars) + '\n\n... (Archivo Muy largo, ver completo en el ZIP)';
                        truncated = true;
                    }

                    // Añadir el encabezado de sección solo una vez
                    if (!addedSectionHeading) {
                        paragraphs.push(sectionHeading);
                        addedSectionHeading = true;
                    }

                    // Heading para permitir plegado/desplegado en Word
                    const headingParagraph = new Paragraph({
                        heading: HeadingLevel.HEADING_3,
                        children: [new TextRun({ text: name, bold: true, size: 20, font: 'Calibri' })]
                    });

                    // Ahora la tabla: dar más espacio a la columna de script (15% / 85%) y fuente ligeramente mayor
                    const tbl = new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                            left: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                            right: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' },
                            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'FFFFFF' }
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        width: { size: 10, type: WidthType.PERCENTAGE },
                                        // Evitar duplicación del nombre: mostrar tipo/etiqueta en vez del nombre repetido
                                        children: [new Paragraph({ children: [new TextRun({ text: 'SQL', bold: true, size: 20, font: 'Calibri' })] })],
                                        margins: { top: 50, bottom: 50, left: 50, right: 50 }
                                    }),
                                    new TableCell({
                                        width: { size: 90, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({ children: [new TextRun({ text: displayed, font: 'Courier New', size: 18 })] })],
                                        margins: { top: 50, bottom: 50, left: 50, right: 50 }
                                    })
                                ]
                            })
                        ]
                    });

                    paragraphs.push(headingParagraph);
                    paragraphs.push(tbl);
                    if (truncated) {
                        paragraphs.push(new Paragraph({ children: [new TextRun({ text: 'Archivo truncado en el documento. Archivo completo incluido en el ZIP.', italics: true, size: 12, font: 'Calibri' })], spacing: { after: 100 } }));
                    }
                } catch (e) {
                    paragraphs.push(new Paragraph({ children: [new TextRun({ text: `Error leyendo ${name}: ${String(e)}`, size: 12, font: 'Calibri' })], spacing: { after: 100 } }));
                }
            }
            return paragraphs;
        }

        // Construye y devuelve el array de Paragraphs para las secciones del documento.
        function buildSections(info) {
            const sections = [];
            let splitAtIndex = null; // index where '4 CHANGES TO APPLY' starts

            // -------------------------
            // Header row (titles centered, more vertical margin)
            new TableRow({
                children: [
                    new TableCell({
                        width: { size: 6, type: WidthType.PERCENTAGE },
                        children: [new Paragraph({ children: [new TextRun({ text: '#', bold: true, size: 20, font: 'Calibri' })], alignment: AlignmentType.CENTER })],
                        shading: { fill: 'c1d59a' },
                        margins: { top: 75, bottom: 75, left: 75, right: 75 }
                    }),
                    new TableCell({
                        width: { size: 30, type: WidthType.PERCENTAGE },
                        children: [new Paragraph({ children: [new TextRun({ text: 'Test Cases Description', bold: true, size: 20, font: 'Calibri' })], alignment: AlignmentType.CENTER })],
                        shading: { fill: 'c1d59a' },
                        margins: { top: 75, bottom: 75, left: 75, right: 75 }
                    }),
                    new TableCell({
                        width: { size: 15, type: WidthType.PERCENTAGE },
                        children: [new Paragraph({ children: [new TextRun({ text: 'Input Data', bold: true, size: 20, font: 'Calibri' })], alignment: AlignmentType.CENTER })],
                        shading: { fill: 'c1d59a' },
                        margins: { top: 75, bottom: 75, left: 75, right: 75 }
                    }),
                    new TableCell({
                        width: { size: 19, type: WidthType.PERCENTAGE },
                        children: [new Paragraph({ children: [new TextRun({ text: 'Step to Follow', bold: true, size: 20, font: 'Calibri' })], alignment: AlignmentType.CENTER })],
                        shading: { fill: 'c1d59a' },
                        margins: { top: 75, bottom: 75, left: 75, right: 75 }
                    }),
                    new TableCell({
                        width: { size: 30, type: WidthType.PERCENTAGE },
                        children: [new Paragraph({ children: [new TextRun({ text: 'Expected Result', bold: true, size: 20, font: 'Calibri' })], alignment: AlignmentType.CENTER })],
                        shading: { fill: 'c1d59a' },
                        margins: { top: 75, bottom: 75, left: 75, right: 75 }
                    })
                ]
            }),
                // 2 GENERAL INFORMATION
                // -------------------------
                sections.push(
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        children: [new TextRun({ text: '2 GENERAL INFORMATION', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                        spacing: { before: 20, after: 20 },
                        shading: { fill: '2E8B57' }
                    })
                );
            sections.push(new Paragraph({ text: ' ', spacing: { after: 200 } }));

            // Subsections for 2.x
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '2.1	SOLUTION SPECIFICATION', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 160 }
                })
            );

            const solution = [
                ['Solution Specification', 'Solution Specification Title', 'Solution Proposal', 'Solution Proposal Title'],
                [data.solman, data.titulo, ' ', ' ']
            ];

            sections.push(makeStyledDynamicTable(solution));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla


            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '2.2 PARTICIPANTS', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 160 }
                })
            );

            const participants = [
                ['Name', 'Rol'],
                [info.usuario.nombre, 'Tester'],
                ['MARTINEZ Benjamin', 'Leader']
            ];

            sections.push(makeStyledDynamicTable(participants));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla


            // -------------------------
            // 3 IMPLEMENTATION PARTICIPANTS
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '3 FUNCIONAL TEST', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 20 },
                    shading: { fill: '2E8B57' }
                })
            );

            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: ' ', font: 'Calibri', size: 20 })],
                    spacing: { after: 120 }
                })
            );

            // Custom table for tests: make second column wider and reduce last column
            const testTable = new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    // Header row
                    new TableRow({
                        children: [
                            new TableCell({
                                width: { size: 6, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ children: [new TextRun({ text: '#', bold: true, size: 20, font: 'Calibri' })], alignment: AlignmentType.CENTER })],
                                shading: { fill: 'c1d59a' },
                                margins: { top: 75, bottom: 75, left: 75, right: 75 }
                            }),
                            new TableCell({
                                width: { size: 30, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ children: [new TextRun({ text: 'Test Cases Description', bold: true, size: 20, font: 'Calibri' })] })],
                                shading: { fill: 'c1d59a' },
                                margins: { top: 75, bottom: 75, left: 75, right: 75 }
                            }),
                            new TableCell({
                                width: { size: 15, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ children: [new TextRun({ text: 'Input Data', bold: true, size: 20, font: 'Calibri' })] })],
                                shading: { fill: 'c1d59a' },
                                margins: { top: 75, bottom: 75, left: 75, right: 75 }
                            }),
                            new TableCell({
                                width: { size: 19, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ children: [new TextRun({ text: 'Step to Follow', bold: true, size: 20, font: 'Calibri' })] })],
                                shading: { fill: 'c1d59a' },
                                margins: { top: 75, bottom: 75, left: 75, right: 75 }
                            }),
                            new TableCell({
                                width: { size: 40, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ children: [new TextRun({ text: 'Expected Result', bold: true, size: 20, font: 'Calibri' })] })],
                                shading: { fill: 'c1d59a' },
                                margins: { top: 75, bottom: 75, left: 75, right: 75 }
                            })
                        ]
                    }),

                    // Data row
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: '1', size: 20, font: 'Calibri' })], alignment: AlignmentType.CENTER })],
                                margins: { top: 75, bottom: 75, left: 75, right: 75 }
                            }),
                            new TableCell({
                                children: [new Paragraph({
                                    children: [
                                        new TextRun({ text: 'Ejecucion de Script en QA.', size: 20, font: 'Calibri' }),
                                        new TextRun({ break: 1 }),
                                        new TextRun({ text: data.procedure, size: 20, bold: true, font: 'Calibri' })
                                    ]
                                })],
                                margins: { top: 75, bottom: 75, left: 75, right: 75 }
                            }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Script', size: 20, font: 'Calibri' })] })], margins: { top: 75, bottom: 75, left: 75, right: 75 } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Ejecutar Script', size: 20, font: 'Calibri' })] })], margins: { top: 75, bottom: 75, left: 75, right: 75 } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: data.resultado || '', size: 20, font: 'Calibri' })] })], margins: { top: 75, bottom: 75, left: 75, right: 75 } })
                        ]
                    })
                ]
            });

            sections.push(testTable);
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla



            // -------------------------
            // 4 CHANGES TO APPLY
            // -------------------------
            // Mark split point: everything from here onward will go to landscape
            splitAtIndex = sections.length;
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '4 INTEGRAL TEST', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 550, after: 160 },
                    shading: { fill: '2E8B57' }
                })
            );
            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: 'Se ejecuta el script en QA.', font: 'Calibri', size: 20 })],
                    spacing: { after: 120 }
                })
            );

            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: data.procedure, font: 'Calibri', size: 20 })],
                    spacing: { after: 120 }
                })
            );

            const testT = [
                ['Test Type', 'X', 'Justification of Excluded Tests / Test Results'],
                ['Data and DB Integrity Test', ' ', ' '],
                ['Business Cycle Test', ' ', ' '],
                ['User Interface Test', ' ', ' '],
                ['Performance Test', ' ', ' '],
                ['Load Test', ' ', ' '],
                ['Stress Test', ' ', ' '],
                ['Volume Test', ' ', ' '],
                ['Access Security and Control', ' ', ' '],
                ['Failure Recovery Test', ' ', ' '],
                ['Configuration Test', ' ', ' '],
                ['Installation Test', ' ', ' '],
                ['Regression Test', ' ', ' ']
            ];
            sections.push(makeStyledDynamicTable(testT));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla

            const other = [
                ['Other Test Included', 'Justification'],
                [' ', ' '],
                [' ', ' '],
                [' ', ' ']
            ];
            sections.push(makeStyledDynamicTable(other));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla


            // -------------------------
            // 5 OBJECTS TO DEPLOY
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '5 OTHER EXECUTED TEST', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 200 },
                    shading: { fill: '2E8B57' }
                })
            );

            const otherTest = [
                ['Test Description', 'Result'],
                ['Data and DB Integrity', 'Databases and database processes should be tested as a separate subsystem. This test should try subsystems without the User Interfaces as data interface. Databases should be inspected to ensure that the data have been stored as expected and that all events in the database were carried out correctly.'],
                ['Functional', [{ text: 'This Test is based on the execution of flows for every individual use case scenario, or the application functionalities and features, using valid and invalid data to verify that:' },
                { text: 'The expected results are obtained when valid data is used', bullet: true },
                { text: 'The appropriate errors and / or alerts are shown when invalid data is used', bullet: true },
                { text: 'All business rules are applied correctly', bullet: true }
                ]
                ],
                ['Business Cycle', 'The Business Cycle Test should emulate the activities performed in a certain period of time. A period should be identified, e.g., a year, and the transactions and activities that would take place in that period should be executed. This includes all daily, weekly and monthly cycles, and date-sensitive events, such as schedules.'],
                ['Performance', 'The Performance Test is a test that measures and evaluates reaction times, transaction rates and other time-sensitive requirements. The goal is to verify that performance requirements have been met. The Performance Test is implemented and executed to analyze and adjust the application performance based on conditions such as work load or hardware configurations.'],
                ['Load', 'The Load Test is a performance test that imposes several work loads on the application to measure and evaluate the application capacity to continue working properly in different scenarios. The Load Test goal is to determine and ensure that the system works correctly beyond the expected maximum work load. In addition, the Load Test evaluates performance features, such as reaction times, transaction fees and other time-sensitive features.'],
                ['Stress', 'The Stress Test is a type of performance test implemented and executed to understand how the system fails as a consequence of limit conditions or beyond the expected tolerances. This typically means low resources or competition for resources. When the application is executed for testing under low resources conditions, it is possible to determine how the application fails, which is not evident under normal conditions. Other defects may arise when the application to be tested competes for shared resources with other applications, such as bandwidth, though some of these tests are usually dealt with in the Functional or Load Tests.'],
                ['Volume', 'The Volume Test imposes very large amounts of data on the application in order to determine if the application reaches some limit where it fails. The Volume Test also identifies the maximum continuous load or volume that the application can handle during a certain period. For instance, if the application is processing a set of database records to generate a report, a volume test would use a large test database and would prove that the software behaved normally and generated the correct report.'],
                ['Access Security and Control', [{ text: 'This Test is based on the execution of flows for every individual use case scenario, or the application functionalities and features, using valid and invalid data to verify that:' },
                { text: 'Security at application level, including access to business data or functionalities.', bullet: true },
                { text: 'Security at system level, including remote logging and access to the system. ', bullet: true },
                { text: ' ',},
                { text: 'Depending on the desired security level, security at user level ensures that actors as restricted to specific functionalities or use cases, or limited in the data available to them. For instance, they can all be allowed to enter data and create new accounts, but only Administrators may delete them. If there is security at data level, the test ensures that Type One Users can see all the information on the customer, including financial data; however, Type Two Users can only see the demographic data corresponding to the same customer. Security at system level ensures that only users with authorized access to the system are able to access applications and only through the right channels.' }
                ]
                ],
                ['Configuration', 'The Configuration Test verifies the operation of the application to be tested, in different software and hardware configurations. In most production environments, hardware specifications for customers, network connections and database servers vary. Customer workstations can have different software packages installed, e.g., applications, drivers, etc. – and, at any given time, different combinations may be active using different resources.'],
                ['Installation', 'The Installation Test has two purposes. The first one is to ensure that the software may be installed under different conditions, such as a new installation, an updated version, or a full or customized installation, under normal and abnormal conditions. Abnormal conditions may be insufficient disk space, lack of privileges to create directories, etc. The second purpose is to verify that, once installed, the application works properly. This usually means executing a set of cases that were developed for the Functional Test.']
            ];
            sections.push(makeStyledTwoColumnTable2(otherTest));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla


            // -------------------------
            // 6 EXECUTION PLAN
            // -------------------------
            // Heading for Appendix I
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '6 APPENDIX I:   TEST TYPES', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24, italics: true })],
                    spacing: { before: 20, after: 200 },
                    shading: { fill: '2E8B57' }
                })
            );

            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: 'Special issues to take into account:', font: 'Calibri', size: 16, italics: true })],
                    spacing: { after: 120 }
                })
            );

            // Bulleted lines (small bullet character so we control size)
            sections.push(new Paragraph({ children: [
                new TextRun({ text: '•\t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Data conversion should be tested between origin and destination to confirm it is complete, accurate, valid and approved.', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '•\t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Ensure that backup/restore/recovery mechanisms are available before the test begins.', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '•\t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Ensure the user participation in data modification tests.', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));


            sections.push(                new Paragraph({
                    children: [new TextRun({ text: ' ', font: 'Calibri', size: 16 })],
                    spacing: { after: 120 }
                })
            );


            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: 'Applicable tests:', font: 'Calibri', size: 16, italics: true })],
                    spacing: { after: 120 }
                })
            );

            // Bulleted lines (small bullet character so we control size)
            sections.push(new Paragraph({ children: [
                new TextRun({ text: '           \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Data and Database Integrity Test', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '           \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Functional Test', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '           \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Business Cycle Test (depending on the impact of the change).', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(                
                new Paragraph({
                children: [new TextRun({ text: ' ', font: 'Calibri', size: 16 })],
                spacing: { after: 120 }
            })
            );


            // -------------------------
            // 7 COMMENTS
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '7 APPENDIX II:  DATA CONVERSION TEST', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 200 },
                    shading: { fill: '2E8B57' }
                })
            );

            
            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: 'Special issues to take into account:', font: 'Calibri', size: 16, italics: true })],
                    spacing: { after: 120 }
                })
            );

             // Bulleted lines (small bullet character so we control size)
            sections.push(new Paragraph({ children: [
                new TextRun({ text: '•\t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Data conversion should be tested between origin and destination to confirm it is complete, accurate, valid and approved.', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '•\t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Ensure that backup/restore/recovery mechanisms are available before the test begins ', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '•\t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Business Cycle Test (depending on the impact of the change).', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));



            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: ' ', font: 'Calibri', size: 16 })],
                    spacing: { after: 120 }
                })
            );

            // -------------------------
            // 8 REVISION HISTORY
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '8 APPENDIX III: INTERFACE TEST', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 200 },
                    shading: { fill: '2E8B57' }
                })
            );

            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: 'Special issues to take into account:', font: 'Calibri', size: 16, italics: true })],
                    spacing: { after: 120 }
                })
            );

            // Bulleted lines (small bullet character so we control size)
            sections.push(new Paragraph({ children: [
                new TextRun({ text: '•\t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Consider if the Failure recovery test level is enough and was included among the tests so that it is possible to analyze how the application behaves in the event of transaction cancellations or incomplete data transmissions.', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '•\t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Ensure that backup/restore/recovery mechanisms are available before the test begins ', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: ' ', font: 'Calibri', size: 16 })],
                    spacing: { after: 120 }
                })
            );

            


            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: '        Applicable tests: ', font: 'Calibri', size: 16 })],
                    spacing: { after: 120 }
                })
            );
            // Bulleted lines (small bullet character so we control size)
            sections.push(new Paragraph({ children: [
                new TextRun({ text: '                \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Data and Database Integrity Test', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '                \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Functional Test', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '                   \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Business Cycle Test (depending on the impact of the change)', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '                   \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Performance Test', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '                   \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Load Test', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));

            sections.push(new Paragraph({ children: [
                new TextRun({ text: '                   \t', font: 'Calibri', size: 14 }),
                new TextRun({ text: 'Volume Test', font: 'Calibri', size: 14, italics: true })
            ], spacing: { after: 100 } }));
            
            sections.push(                
                new Paragraph({
                children: [new TextRun({ text: ' ', font: 'Calibri', size: 16 })],
                spacing: { after: 120 }
            })
            );

            // -------------------------
            // 9 REVISION HISTORY
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '9 APPENDIX IV: REVISION HISTORY', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 200 },
                    shading: { fill: '2E8B57' }
                })
            );

            const revision = [
                ['Date', 'Version', 'Description', 'Approved / Revised by', 'Author'],
                [new Date().toLocaleDateString('es-ES'), '1.0', data.descripcion, 'Martinez Benjamin', info.usuario.nombre]
            ];
            sections.push(makeStyledDynamicTable(revision));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla



            // If no split point found, set to full length
            if (splitAtIndex === null) splitAtIndex = sections.length;

            const portrait = sections.slice(0, splitAtIndex);
            const landscape = sections.slice(splitAtIndex);
            return { portrait, landscape };
        }

        // Instancias reutilizables
        const headerTable = makeHeaderTable(imageBuffer);
        const horizontalLine = makeHorizontalLine();
        const secondPageHeader = makeSecondPageHeader(data);
        const horizontalLine2 = makeHorizontalLine();

        const split = buildSections(data);

        // Si hay archivos adjuntos (respaldos), convertirlos en Paragraphs e insertarlos en el punto 4 (inicio de landscape)
        let attachmentParagraphs = [];
        try {
            attachmentParagraphs = await makeAttachmentParagraphs(data.respaldos);
            if (attachmentParagraphs.length > 0) {
                const insertIndex = Math.min(3, split.landscape.length);
                split.landscape.splice(insertIndex, 0, ...attachmentParagraphs);

            }
        } catch (e) {
            console.warn('No se pudieron procesar los archivos adjuntos:', e);
        }

        const doc = new Document({
            sections: [
                // Primera sección (primera página)
                {
                    properties: {},
                    headers: {
                        default: new Header({
                            children: [headerTable, horizontalLine]
                        })
                    },
                    children: [
                        // Espaciado superior para centrar verticalmente
                        new Paragraph({
                            children: [new TextRun({ text: ' ', size: 1 })],
                            spacing: { before: 3700 }
                        }),

                        // Título principal
                        new Paragraph({
                            children: [new TextRun({ text: `${data.solman}, ${data.titulo}`, bold: true, italics: true, size: 44, font: 'Calibri' })],
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 10 }
                        }),

                        // Línea horizontal grande
                        new Paragraph({
                            children: [new TextRun({ text: '___________________________________________', size: 40, font: 'Calibri', bold: true })],
                            alignment: AlignmentType.CENTER,
                            spacing: { before: 0, after: 100 }
                        }),

                        // Subtítulos
                        new Paragraph({
                            children: [new TextRun({ text: 'Test Evidence', size: 48, font: 'Calibri' })],
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 1100 }
                        }),

                        new Paragraph({
                            children: [new TextRun({ text: 'Release Number', bold: true, size: 32, font: 'Calibri' })],
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 0 }
                        }),

                        new Paragraph({
                            children: [new TextRun({ text: `${data.solman}`, bold: true, size: 32, font: 'Calibri' })],
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 1100 }
                        }),

                        new Paragraph({
                            children: [new TextRun({ text: 'Version 1.1', size: 32, bold: true, font: 'Calibri' })],
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 1000 }
                        }),

                        new Paragraph({ children: [new PageBreak()] })
                    ]
                },

                // TOC
                {
                    properties: {},
                    headers: {
                        default: new Header({ children: [secondPageHeader, horizontalLine2] })
                    },
                    children: [
                        new Paragraph({
                            heading: HeadingLevel.HEADING_1,
                            children: [new TextRun({ text: 'TABLE OF CONTENTS', bold: true, color: 'FFFFFF', size: 24, font: 'Calibri' })],
                            spacing: { after: 200 },
                            shading: { fill: '2E8B57' }
                        }),
                        // Índice manual: párrafos con líder de puntos y número al final
                        ...makeManualTOC(
                            ['1 OVERVIEW', '2 GENERAL INFORMATION', '3 IMPLEMENTATION PARTICIPANTS', '4 CHANGES TO APPLY', '5 OBJECTS TO DEPLOY', '6 EXECUTION PLAN', '7 COMMENTS', '8 REVISION HISTORY'],
                            // Números fijos (puedes ajustarlos según el documento final)
                            ['1', '2', '3', '4', '5', '6', '7', '8']
                        ),
                        // Salto de página para empezar las secciones en la siguiente
                        new Paragraph({ children: [new PageBreak()] })
                    ]
                },

                // Contenido (todas las secciones en portrait)
                {
                    properties: {},
                    headers: { default: new Header({ children: [secondPageHeader, horizontalLine2] }) },
                    children: [...split.portrait, ...split.landscape]
                }
            ]
        });
        return await Packer.toBlob(doc);
    } catch (error) {
        console.error("Error generando Release Plan:", error);
        throw error;
    }
}
