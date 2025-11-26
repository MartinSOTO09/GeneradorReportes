async function generarReleasePlan_linux(data) {
    if (typeof docx === 'undefined') {
        throw new Error("La librería docx no está cargada");
    }

    const { Document, Packer, Paragraph, TextRun, AlignmentType, Header, Table, TableRow, TableCell, WidthType, ImageRun, BorderStyle, PageBreak, TableOfContents, HeadingLevel, UnderlineType } = docx;

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
        function makeStyledTwoColumnTable(rowsData) {
            return new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: rowsData.map(([label, value]) =>
                    new TableRow({
                        children: [
                            // Columna izquierda (fondo verde)
                            new TableCell({
                                width: { size: 28, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: label,
                                                font: 'Calibri',
                                                size: 20,
                                                bold: true,
                                                color: '000000'
                                            })
                                        ],
                                        alignment: AlignmentType.LEFT
                                    })
                                ],
                                shading: { fill: 'c1d59a' },
                                margins: { top: 50, bottom: 50, left: 100, right: 100 }
                            }),

                            // Columna derecha (blanca)
                            new TableCell({
                                width: { size: 72, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: value,
                                                font: 'Calibri',
                                                size: 20,
                                                color: '000000'
                                            })
                                        ],
                                        alignment: AlignmentType.LEFT
                                    })
                                ],
                                margins: { top: 50, bottom: 50, left: 200, right: 200 }
                            })
                        ]
                    })
                )
            });
        }

        /* Tabla dinamica con N filas y N columnas */
        function makeStyledDynamicTable(rowsData) {
            return new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: rowsData.map((row, rowIndex) =>
                    new TableRow({
                        children: row.map((cellContent) => {
                            // Convertir a array de objetos si es string simple
                            let contentArray;
                            if (typeof cellContent === 'string') {
                                // Dividir por salto de línea
                                contentArray = cellContent.split('\n').map(line => ({ text: line }));
                            } else if (Array.isArray(cellContent)) {
                                // Suponemos que ya es un array de objetos
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

        function makeCustomEightColumnTable(rowsDataOrWithWidths, colWidths, centerRows) {
            // Permite dos formas de uso:
            // 1) makeCustomEightColumnTable([ [w1..w8], row1, row2, ... ])  -> un solo arreglo, el primero son los anchos
            // 2) makeCustomEightColumnTable(rowsData, colWidths)            -> anchos por parámetro separado
            let rowsData = Array.isArray(rowsDataOrWithWidths) ? [...rowsDataOrWithWidths] : [];
            let widths = [12.5, 12.5, 12.5, 12.5, 12.5, 12.5, 12.5, 12.5];

            if (rowsData.length && Array.isArray(rowsData[0]) && rowsData[0].length === 8 && rowsData[0].every(n => typeof n === 'number')) {
                widths = rowsData.shift();
            } else if (Array.isArray(colWidths) && colWidths.length === 8) {
                widths = colWidths;
            }

            // Helper: insertar puntos de quiebre suaves para URLs o palabras largas
            const insertSoftWraps = (s) => {
                if (typeof s !== 'string') return s;
                // Inserta zero-width space después de delimitadores comunes para permitir corte de línea
                return s
                    .replaceAll('/', '/\u200B')
                    .replaceAll('?', '?\u200B')
                    .replaceAll('&', '&\u200B')
                    .replaceAll('=', '=\u200B')
                    .replaceAll('-', '-\u200B')
                    .replaceAll('_', '_\u200B');
            };

            return new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                layout: docx.TableLayoutType.FIXED,
                rows: rowsData.map((row, rowIndex) =>
                    new TableRow({
                        children: row.slice(0, 8).map((cellContent, idx) => {
                            let contentArray;
                            if (typeof cellContent === 'string') {
                                contentArray = cellContent.split('\n').map(line => ({ text: insertSoftWraps(line) }));
                            } else if (Array.isArray(cellContent)) {
                                // Aplicar soft wraps a cada item.text si es string
                                contentArray = cellContent.map(it => ({
                                    ...it,
                                    text: typeof it.text === 'string' ? insertSoftWraps(it.text) : it.text
                                }));
                            } else if (cellContent && typeof cellContent === 'object') {
                                // Permitir un único objeto { text, ... }
                                const it = cellContent;
                                contentArray = [{
                                    ...it,
                                    text: typeof it.text === 'string' ? insertSoftWraps(it.text) : it.text
                                }];
                            } else {
                                contentArray = [{ text: '' }];
                            }

                            const paragraph = new Paragraph({
                                children: contentArray.flatMap((item, i) => {
                                    const tr = new TextRun({
                                        text: item.text,
                                        font: item.font || 'Calibri',
                                        size: item.size || 20,
                                        bold: item.bold || (rowIndex === 0),
                                        italic: item.italic || false,
                                        color: item.color || '000000',
                                        highlight: item.highlight ? String(item.highlight).toLowerCase() : undefined,
                                        underline: item.underline
                                    });
                                    return i < contentArray.length - 1 ? [tr, new TextRun({ break: 1 })] : [tr];
                                }),
                                alignment: centerRows ? AlignmentType.CENTER : (rowIndex === 0 ? AlignmentType.CENTER : AlignmentType.LEFT)
                            });

                            return new TableCell({
                                width: { size: widths[idx] || 0, type: WidthType.PERCENTAGE },
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



        // Construye y devuelve el array de Paragraphs para las secciones del documento.
        function buildSections(info) {
            const sections = [];
            let splitAtIndex = null; // index where '4 CHANGES TO APPLY' starts

            // -------------------------
            // 1 OVERVIEW
            // -------------------------

            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [
                        new TextRun({ text: '1 OVERVIEW', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })
                    ],
                    spacing: { before: 20, after: 20 },
                    shading: { fill: '2E8B57' }
                })
            );


            sections.push(new Paragraph({ text: ' ', spacing: { after: 80 } }));


            // Subsections for 1.x
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [
                        new TextRun({ text: '1.1 Objective', font: 'Calibri', size: 24, bold: true, color: '000000' })
                    ],
                    spacing: { after: 120 }
                })
            );

            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: data.objetivo, font: 'Calibri', size: 20, color: '4F81BD', italics: true })],
                    spacing: { after: 200 }
                })
            );

            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [
                        new TextRun({ text: '1.2 References', font: 'Calibri', size: 24, bold: true, color: '000000' })
                    ],
                    spacing: { after: 120 }
                })
            );

            // -------------------------
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


            const participants = [
                ['Release Number', data.solman],
                ['Release Type', 'Minor'],
                ['Solution Specification number ', ''],
                ['System Name', data.sistemas],
                ['Package/Label Number in the SCM tool', 'Change Manager'],
                ['USD Number', data.ticket]
            ];

            sections.push(makeStyledTwoColumnTable(participants));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla

            // -------------------------
            // 3 IMPLEMENTATION PARTICIPANTS
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '3 IMPLEMENTATION PARTICIPANTS', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 20 },
                    shading: { fill: '2E8B57' }
                })
            );
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } }));

            const rolesTableData = [
                ['Rol', 'Responsibilities - Activities to Carry Out'],
                ['Technical Project Leader', 'IT L2 Industrial Support Coordination <IT.L2.INDSC@tenaris.com>'],
                ['User Leader', ''],
                ['Project Leader', 'IT L2 Industrial Support Coordination <IT.L2.INDSC@tenaris.com>'],
                ['Technical Analyst', info.usuario.nombre],
                ['Technologist', 'Grupo de Implementadores']
            ];

            sections.push(makeStyledDynamicTable(rolesTableData));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla

            // -------------------------
            // 4 CHANGES TO APPLY
            // -------------------------
            // Mark split point: everything from here onward will go to landscape
            splitAtIndex = sections.length;
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '4 CHANGES TO APPLY', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 550, after: 160 },
                    shading: { fill: '2E8B57' }
                })
            );

            // Subsections for 4.x
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '4.1 AFFECTED ENVIRONMENTS (CHECKMARK):', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 160 }
                })
            );


            // Function to create a TextRun for a checkbox option
            const makeCheckboxRun = (text, isChecked) => {
                // Crear dos TextRuns separados: uno para el texto y otro para la marca
                const textRun = new TextRun({
                    text: text,
                    font: 'Calibri',
                    size: 20,
                    color: '000000',
                    highlight: isChecked ? 'yellow' : undefined
                });

                const checkboxRun = new TextRun({
                    text: isChecked ? ' [X]' : ' [ ]',
                    font: 'Calibri',
                    size: 20,
                    color: '000000',
                    highlight: isChecked ? 'yellow' : undefined
                });

                return [textRun, checkboxRun];
            };

            // Create array of TextRuns based on checkmark data from form
            const checkboxes = [
                { text: 'AS400', field: 'AS400' },
                { text: 'Web', field: 'Web' },
                { text: 'Oracle', field: 'Oracle' },
                { text: 'SAP', field: 'SAP' },
                { text: 'ETL', field: 'ETL' },
                { text: 'HOST', field: 'HOST' },
                { text: 'DWT', field: 'DWT' },
                { text: 'BATCH WIN', field: 'BATCH_WIN' }
            ];

            console.log("Checkmarks recibidos:", data.checkmarks); // Debug

            const checkmarkRuns = [];
            checkboxes.forEach((box, i) => {
                // Add space between checkboxes except for the first one
                if (i > 0) {
                    checkmarkRuns.push(new TextRun({ text: '   ', font: 'Calibri', size: 24 }));
                }

                // Verificar si el checkbox está marcado (usando el valor exacto como está en el HTML)
                const isChecked = box.text === 'HOST' ? true : (Array.isArray(data.checkmarks) && data.checkmarks.includes(box.text));
                console.log(`Checkbox ${box.text}: ${isChecked}`); // Debug

                // Crear los TextRuns para el texto y el checkbox
                const runs = makeCheckboxRun(box.text, isChecked);
                checkmarkRuns.push(...runs);
            });

            sections.push(
                new Paragraph({
                    children: checkmarkRuns,
                    spacing: { after: 160 }
                })
            );

            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: 'Other (specify):', font: 'Calibri', size: 20 })],
                    spacing: { after: 160 }
                })
            );

            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '4.2 AFFECTED SYSTEMS', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 160 }
                })
            );

            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '4.3 SCENARIOS WHERE CHANGES WILL BE APPLIED (CHECKMARK):', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 160 }
                })
            );

            const scenarios = [
                ['', 'Scenario', 'Description', 'Environments'],
                ['[X]', 'Net Intranet', '.Net applications hosted in the new infrastructure.', '.Net Development, Citrix, .Net Local Test, .Net Integral Test, .Net Production '],
                ['[ ]', 'DNA old infrastructure', 'DNA applications still not migrated to the new infrastructure.', 'DNA Development, DNA Pre-production, DNA Production '],
                ['[ ]', 'DNA new infrastructure Intra2', 'DNA management application (TMC) migrated to the new infrastructure.', 'DNA Local Test, DNA Integral Test, DNA Production Intra2'],
                ['[ ]', 'DNA new infrastructure Intra3', 'DNA applications migrated to the new infrastructure (except TMC).', 'DNA Local Test, DNA Integral Test, DNA Production Intra3'],
                ['[ ]', '.Net Internet', '.Net applications hosted in the new Internet infrastructure.', '.Net Development, Citrix, .Net Local Test, .Net Integral Test, .Net Internet Integral Test, .Net Internet Production'],
                ['[ ]', 'DNA Internet', 'Batch processes.', 'DNA Local Test, DNA Integral Test, DNA Internet Production'],
                ['[ ]', 'Batch', 'Mainframe applications.', '.Net Batch QA, .Net Batch Production'],
                ['[X]', 'Oracle Intranet', 'Intranet transaction Data Bases.', 'Oracle Development, Oracle QA, Oracle Operative'],
                ['[ ]', 'Oracle Internet', 'Internet transaction Data Bases.', 'Oracle Development, Oracle QA, Oracle Operative, Oracle Internet Operative']
            ];
            sections.push(makeStyledDynamicTable(scenarios));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 370 } })); // Espacio después de la tabla

            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '4.4 HARDWARE AND SOFTWARE CONFIGURATIONS', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 140 }
                })
            );

            const hardwareSoftware = [
                ['Configuration ID', 'HW/SW type', 'Element', 'Secenario', 'Configuration', 'USD Number'],
                [
                    'SW01',
                    'SW',
                    '',
                    '',
                    '',
                    ''
                ]

            ];

            // Agregar filas para casos adicionales de Linux (SW02..SW0N)
            try {
                const linuxExtras = Array.isArray(data.extra_cases) ? data.extra_cases.filter(c => c && c.type === 'linux') : [];
                linuxExtras.forEach((c, idx) => {
                    const swId = `SW0${idx + 2}`;
                    hardwareSoftware.push([
                        swId,
                        'SW',
                        '',
                        '',
                        '',
                        ''
                    ]);
                });
            } catch(_) {}

            sections.push(makeStyledDynamicTable(hardwareSoftware));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla


            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '4.5 SOFTWARE INVENTORY', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 120 }
                })
            );

            const softwareInventory = [
                ['ID Component', 'Software Component or Objects', 'Type', 'Comments'],
                ['', '', '', ''],
                ['', '', '', '']
            ];

            sections.push(makeStyledDynamicTable(softwareInventory));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla

            // -------------------------
            // 5 OBJECTS TO DEPLOY
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '5 OBJECTS TO DEPLOY', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 20 },
                    shading: { fill: '2E8B57' }
                })
            );

            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '5.1 LIST OF MAINFRAME OBJECTS', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 120 }
                })
            );

            const mainframeObjects = [
                ['Type of Object', 'Object name/s ', 'Version/s', 'Location', 'Object date and time', 'Observations'],
                ['Global', '', '', '', '', ''],
                ['Local', '', '', '', '', ''],
                ['Subroutines', '', '', '', '', ''],
                ['Maps', '', '', '', '', ''],
                ['Subprograms', '', '', '', '', ''],
                ['Copycode', '', '', '', '', ''],
                ['Programs', '', '', '', '', ''],
                ['Help Routine', '', '', '', '', ''],
                ['Parameters', '', '', '', '', ''],
                ['COBOL I', '', '', '', '', ''],
                ['COBOL II', '', '', '', '', '']
            ];
            sections.push(makeStyledDynamicTable(mainframeObjects));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla

            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '5.2 LIST OF ORACLE OBJECTS', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 120 }
                })
            );

            const oracleObjects = [
                ['Type of Object', 'Object name/s ', 'Version/s', 'Source database', 'Source schema', 'Target database', 'Target schema', 'Observations'],
                ['Stored Procedure', '', '', '', '', '', '', ''],
                ['Index', '', '', '', '', '', '', ''],
                ['Trigger', '', '', '', '', '', '', ''],
                ['Sequence', '', '', '', '', '', '', '']
            ];
            sections.push(makeStyledDynamicTable(oracleObjects));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla


            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '5.3 LIST OF JOBS', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 120 }
                })
            );

            const jobs = [
                ['Processing platform', 'Name/s', 'Version/es', 'Source location', 'Target location', 'Object date and time', 'Observations'],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', ''],
                ['', '', '', '', '', '', '']
            ];
            sections.push(makeStyledDynamicTable(jobs));
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } })); // Espacio después de la tabla


            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '5.4 RELATED ADABAS DATA MODEL', font: 'Calibri', size: 24, bold: true, color: '000000' })],
                    spacing: { after: 120 }
                })
            );

            sections.push(new Paragraph({ children: [new PageBreak()] }));
            // -------------------------
            // 6 EXECUTION PLAN
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '6 EXECUTION PLAN', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 200 },
                    shading: { fill: '2E8B57' }
                })
            );

            const exampleHeaders = [
                [{ text: 'Action – Activity', bold: true }],
                [{ text: 'Sequence Number', bold: true }],
                [{ text: 'Responsible Person', bold: true }],
                [{ text: 'Suggested Implementation Time', bold: true }],
                [{ text: 'Dependent components', bold: true }],
                [{ text: 'ITDS Mandatory Considerations', bold: true }],
                [{ text: 'ITDS Suggested Considerations', bold: true }],
                [{ text: 'Status', bold: true }]
            ];
            const exampleRow = [[{ text: 'Autorizar la solicitud' },
            { text: ' ' },
            { text: 'Tamsa: ' + data.solicitud + ' - ' + data.nombre_solicitud, bold: true },
            { text: ' ' },
            { text: 'desde el comando APT - Soporte' }
            ], '1',
            [{ text: 'MARTINEZ R. Benjamin F.' },
            { text: ' ' },
            { text: 'TENARIS IT ' },
            { text: 'bfmartinezr@tenaris.com' },
            { text: ' ' },
            { text: 'DIAZ Angelica TENARIS' },
            { text: 'ADIAZ@tenaris.com' }
            ], 'ASAP', ' ', ' ', ' ', { text: '[Pending]', highlight: 'yellow', underline: UnderlineType.SINGLE }];


            const exampleRow2 = [[{ text: 'Ejecución de la solicitud ' },
            { text: 'Tamsa: ' + data.solicitud + ' - ' + data.nombre_solicitud, bold: true }
            ], '2', [{ text: info.usuario.nombre },
            { text: info.usuario.email }
            ], 'ASAP', ' ', ' ', ' ', { text: '[Pending]', highlight: 'yellow', underline: UnderlineType.SINGLE }];


            // Construir filas dinámicas agregando un renglón por cada caso adicional de Linux (basado en ejecución de solicitud)
            const rowsRP6 = [exampleHeaders, exampleRow, exampleRow2];
            try {
                const linuxExtras = Array.isArray(data.extra_cases) ? data.extra_cases.filter(c => c && c.type === 'linux') : [];
                linuxExtras.forEach((c, idx) => {
                    const seq = String(idx + 3); // después de 1 y 2
                    const row = [[
                        { text: 'Ejecución de la solicitud ' },
                        { text: 'Tamsa: ' + (c.solicitud || '') + ' - ' + (c.nombre_solicitud || ''), bold: true }
                    ], seq, [{ text: info.usuario.nombre }, { text: info.usuario.email }], 'ASAP', ' ', ' ', ' ', { text: '[Pending]', highlight: 'yellow', underline: UnderlineType.SINGLE }];
                    rowsRP6.push(row);
                });
            } catch(_) {}

            const exampleTable = makeCustomEightColumnTable([[24, 9, 21, 11, 9, 9, 8, 10], ...rowsRP6], undefined, true);
            sections.push(exampleTable);
            sections.push(new Paragraph({ text: ' ', spacing: { after: 120 } }));

            // -------------------------
            // 7 COMMENTS
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '7 COMMENTS', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 200 },
                    shading: { fill: '2E8B57' }
                })
            );

            sections.push(
                new Paragraph({
                    children: [new TextRun({ text: ' ', font: 'Calibri', size: 20 })],
                    spacing: { after: 120 }
                })
            );

            // -------------------------
            // 8 REVISION HISTORY
            // -------------------------
            sections.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({ text: '8 REVISION HISTORY', bold: true, color: 'FFFFFF', font: 'Calibri', size: 24 })],
                    spacing: { before: 20, after: 200 },
                    shading: { fill: '2E8B57' }
                })
            );

            const revisionHistory = [
                ['Date', 'Version', 'Description', 'Approved / Revised by', 'Author'],
                [new Date().toLocaleDateString('es-ES'), '1.0', 'Creación del documento', 'Martinez Benjamin', info.usuario.nombre]
            ];
            sections.push(makeStyledDynamicTable(revisionHistory));
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
                            children: [new TextRun({ text: 'Release Plan', size: 48, font: 'Calibri' })],
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

                // Contenido (secciones reales)
                //  - Sección en portrait (contenido hasta '4 CHANGES TO APPLY')
                {
                    properties: {},
                    headers: { default: new Header({ children: [secondPageHeader, horizontalLine2] }) },
                    children: split.portrait
                },
                // Sección en landscape (a partir de '4 CHANGES TO APPLY')
                {
                    properties: {
                        page: {
                            size: {
                                orientation: docx.PageOrientation.LANDSCAPE
                            }
                        }
                    },
                    headers: { default: new Header({ children: [secondPageHeader, horizontalLine2] }) },
                    children: [...split.landscape]
                }
            ]
        });
        return await Packer.toBlob(doc);
    } catch (error) {
        console.error("Error generando Release Plan:", error);
        throw error;
    }
}
