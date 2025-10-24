document.addEventListener("DOMContentLoaded", async () => {
    const form = document.getElementById("reporteForm");
    const usuarioSelect = document.getElementById("usuario");
    // Selected files from dropzone (use this before falling back to input.files)
    let selectedFiles = null;
    const dropzone = document.querySelector('.dropzone');
    const respInput = document.getElementById('respaldos');
    // Gracefully wire dropzone interactions if present
    if (dropzone && respInput) {
        const dzSelect = dropzone.querySelector('.dz-select');
        const dzText = dropzone.querySelector('.dz-text');

        const updateDropzoneFiles = (files) => {
            selectedFiles = files && files.length ? files : null;
            if (!selectedFiles) {
                if (dzText) dzText.textContent = 'Arrastre o';
            } else {
                const names = Array.from(selectedFiles).map(f => f.name).join(', ');
                if (dzText) dzText.textContent = `${selectedFiles.length} archivo(s): ${names}`;
            }
        };

        // Click the hidden input when user presses the select button
        if (dzSelect) dzSelect.addEventListener('click', () => respInput.click());

        // When the real input changes, capture the files
        respInput.addEventListener('change', (e) => {
            updateDropzoneFiles(e.target.files);
        });

        // Drag & drop handlers
        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.classList.add('dragover');
        });
        dropzone.addEventListener('dragleave', (e) => {
            dropzone.classList.remove('dragover');
        });
        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.classList.remove('dragover');
            if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
                updateDropzoneFiles(e.dataTransfer.files);
            }
        });
    }

    // Cargar usuarios desde el archivo JSON
    try {
    // ruta al JSON (archivo existente: json/usurios.json)
    const response = await fetch('json/usurios.json');
        const usuarios = await response.json();
        
        usuarios.forEach(usuario => {
            const option = document.createElement('option');
            option.value = usuario.id;
            option.textContent = `${usuario.nombre} (${usuario.email})`;
            option.dataset.nombre = usuario.nombre;
            option.dataset.email = usuario.email;
            usuarioSelect.appendChild(option);
        });
        // Auto-populate name/email when user selected
        usuarioSelect.addEventListener('change', () => {
            const sel = usuarioSelect.options[usuarioSelect.selectedIndex];
            const nombreInput = document.getElementById('usuario_nombre');
            const emailInput = document.getElementById('usuario_email');
            if (!sel || !sel.value) {
                if (nombreInput) nombreInput.value = '';
                if (emailInput) emailInput.value = '';
                return;
            }
            if (nombreInput) nombreInput.value = sel.dataset.nombre || '';
            if (emailInput) emailInput.value = sel.dataset.email || '';
        });
    
        // DB type selector: affect UI and will be passed to generators
        const dbTypeSelect = document.getElementById('db_type');
        if (dbTypeSelect) {
            // Add a body class corresponding to the selection so CSS or other scripts can react
            const applyDbBodyClass = (val) => {
                document.body.classList.remove('db-oracle', 'db-postgres', 'db-linux');
                document.body.classList.add(`db-${val}`);
            };
            applyDbBodyClass(dbTypeSelect.value || 'oracle');
            const applyFieldsForDb = (val) => {
                const hideForExternal = (val === 'postgres' || val === 'linux');
                // Elements to hide: bases, esquemas, and checkmark row
                const toHideSelectors = ['#base_origen', '#esquema_origen', '#base_destino', '#esquemas_destino'];
                toHideSelectors.forEach(sel => {
                    const el = document.querySelector(sel);
                    if (el) el.closest && el.closest('div') ? el.closest('div').style.display = (hideForExternal ? 'none' : '') : el.style.display = (hideForExternal ? 'none' : '');
                });

                // Checkmark row container
                const chkRow = document.querySelector('.checkbox-row');
                if (chkRow) chkRow.style.display = (hideForExternal ? 'none' : 'flex');

                // Convert procedure input to a URL-like input (or back to text) and change placeholder/label
                const proc = document.getElementById('procedure');
                if (proc) {
                    try {
                        if (hideForExternal) {
                            // change to url type to hint link usage
                            proc.type = 'url';
                            proc.placeholder = 'Pega aquí el enlace o comando (ej: https://... o ssh://...)';
                            // add a small helper note if not present
                            let helper = proc.nextElementSibling;
                            if (!helper || !helper.classList || !helper.classList.contains('proc-helper')) {
                                helper = document.createElement('div');
                                helper.className = 'proc-helper';
                                helper.style.fontSize = '12px';
                                helper.style.color = '#666';
                                helper.style.marginTop = '4px';
                                proc.parentNode.appendChild(helper);
                            }
                            helper.textContent = 'Para Postgres/Linux use un enlace o comando en lugar de un nombre de procedimiento.';
                        } else {
                            proc.type = 'text';
                            proc.placeholder = 'Nombre del procedimiento';
                            const helper = proc.parentNode.querySelector('.proc-helper');
                            if (helper) helper.remove();
                        }
                    } catch (e) {
                        // ignore DOM quirks
                    }
                }

                // Also toggle bases/schemas on their labels in the form grid for better UX
                const lbls = document.querySelectorAll('label');
                lbls.forEach(l => {
                    if (hideForExternal) {
                        if (/Base Origen|Esquema Origen|Base Destino|Esquema Destino/i.test(l.textContent)) l.style.opacity = '0.6';
                    } else {
                        if (/Base Origen|Esquema Origen|Base Destino|Esquema Destino/i.test(l.textContent)) l.style.opacity = '';
                    }
                });

                // Inline handling for Postgres: inject compact controls into the existing Información General (.info-left)
                const infoLeft = document.querySelector('.info-left');
                // First, handle visibility of original fields
                const fieldsToHandle = ['solman', 'ticket', 'titulo', 'procedure', 'sistemas'];
                fieldsToHandle.forEach(id => {
                    const orig = document.getElementById(id);
                    if (orig) {
                        const container = orig.closest('.form-group') || orig.parentElement;
                        if (container) {
                            container.style.display = val === 'postgres' ? 'none' : '';
                        }
                    }
                });

                if (val === 'postgres') {
                    if (infoLeft) {
                        // Remove any previous injected container
                        const existing = document.getElementById('postgres-compact-inline');
                        if (existing) existing.remove();

                        const container = document.createElement('div');
                        container.id = 'postgres-compact-inline';
                        container.style.display = 'grid';
                        container.style.gridTemplateColumns = '200px 200px 1fr';
                        container.style.gap = '0.6rem';
                        container.style.marginTop = '0.8rem';

                        // Desired order: Solman, Ticket, Titulo, Link (procedure), Sistema, Checkmark (only Postgres option)
                        const fields = [
                            { id: 'solman', label: 'Solman' },
                            { id: 'ticket', label: 'Ticket' },
                            { id: 'titulo', label: 'Título' },
                            { id: 'procedure', label: 'Link' },
                            { id: 'sistemas', label: 'Sistema' }
                        ];

                        fields.forEach(f => {
                            const orig = document.getElementById(f.id);
                            const wrapper = document.createElement('div');
                            const lab = document.createElement('label');
                            lab.textContent = f.label;
                            wrapper.appendChild(lab);
                            if (orig) {
                                const clone = orig.cloneNode(true);
                                clone.id = f.id + '_pg_clone';
                                clone.className = 'form-control';
                                // clear required attribute on clone to avoid duplication issues
                                clone.required = false;
                                if (f.id === 'procedure') {
                                    clone.placeholder = 'Pega aquí el enlace o comando';
                                }
                                wrapper.appendChild(clone);
                            } else {
                                const inp = document.createElement('input');
                                inp.type = 'text';
                                inp.id = f.id + '_pg_clone';
                                inp.className = 'form-control';
                                wrapper.appendChild(inp);
                            }
                            // If this is the sistemas field, attach the Postgres checkbox next to it
                            // (no inline checkbox inserted here; a dedicated grid cell will be appended after the fields loop)

                            container.appendChild(wrapper);
                        });

                        // Add a final grid cell to the container for the Postgres Check Mark (so it sits in the same row as Link and Sistema)
                        const chkWrap = document.createElement('div');
                        const chkTitle = document.createElement('label');
                        chkTitle.textContent = 'Check Mark';
                        chkTitle.style.fontWeight = '600';
                        chkTitle.style.display = 'block';
                        chkWrap.appendChild(chkTitle);
                        const chkLabel = document.createElement('label');
                        chkLabel.innerHTML = '<input style="margin-right:6px;" type="checkbox" value="Postgres" id="chk_postgres" checked>Postgres';
                        chkWrap.appendChild(chkLabel);
                        container.appendChild(chkWrap);

                        // Insert the container near the top of .info-left
                        infoLeft.insertBefore(container, infoLeft.firstChild);
                    }
                } else {
                    // Show original fields and checkmark row when not in postgres mode
                    const fieldsToHandle = ['solman', 'ticket', 'titulo', 'procedure', 'sistemas'];
                    fieldsToHandle.forEach(id => {
                        const orig = document.getElementById(id);
                        if (orig) {
                            const container = orig.closest('.form-group') || orig.parentElement;
                            if (container) {
                                container.style.display = '';
                            }
                        }
                    });
                    
                    // Show the original checkmark row and ensure it has a preceding 'Check Mark' label
                    const originalCheckRow = document.querySelector('.checkbox-row');
                    if (originalCheckRow) {
                        originalCheckRow.style.display = 'flex';
                        // If there is no label immediately before the checkbox-row, insert one
                        const prev = originalCheckRow.previousElementSibling;
                        if (!prev || prev.textContent.trim() !== 'Check Mark') {
                            const lbl = document.createElement('label');
                            lbl.textContent = 'Check Mark';
                            // insert the label just before the checkbox-row
                            originalCheckRow.parentElement.insertBefore(lbl, originalCheckRow);
                        }
                    }
                    // remove inline container if present
                    const existingInline = document.getElementById('postgres-compact-inline');
                    if (existingInline) existingInline.remove();
                }
            };

            // initial apply
            applyFieldsForDb(dbTypeSelect.value || 'oracle');

            dbTypeSelect.addEventListener('change', (e) => {
                applyDbBodyClass(e.target.value || 'oracle');
                applyFieldsForDb(e.target.value || 'oracle');
            });
        }
    } catch (error) {
        console.error('Error cargando usuarios:', error);
        alert('Error cargando la lista de usuarios');
    }

    /**
     * Animates clearing of form fields while preserving the selected user profile.
     * - Staggers fade-out of each field, clears the value, then fades it back in.
     * - Keeps `#usuario` selected and restores derived name/email fields.
     */
    function animateClearPreserveUser() {
        try {
            const usuarioVal = usuarioSelect.value;
            const inputs = Array.from(form.querySelectorAll('input, textarea, select'));
            const toClear = inputs.filter(el => el.id !== 'usuario');

            // Apply transition styles
            toClear.forEach(el => {
                el.style.transition = 'opacity 260ms ease, transform 260ms ease';
                el.style.willChange = 'opacity, transform';
            });

            // Staggered animation: fade out -> clear -> fade in
            toClear.forEach((el, i) => {
                const delay = i * 90; // ms
                setTimeout(() => {
                    el.style.opacity = '0';
                    el.style.transform = 'translateY(-6px)';
                    setTimeout(() => {
                        // Clear value according to element type
                        if (el.tagName.toLowerCase() === 'select') el.selectedIndex = 0;
                        else if (el.type === 'checkbox' || el.type === 'radio') el.checked = false;
                        else if (el.type === 'file') el.value = '';
                        else el.value = '';

                        // Fade back in
                        el.style.transform = 'translateY(0)';
                        el.style.opacity = '1';
                    }, 260);
                }, delay);
            });

            // After animation completes, restore usuario and dropzone state
            const totalDelay = toClear.length * 90 + 320;
            setTimeout(() => {
                usuarioSelect.value = usuarioVal;
                const sel = usuarioSelect.options[usuarioSelect.selectedIndex];
                const nombreInput = document.getElementById('usuario_nombre');
                const emailInput = document.getElementById('usuario_email');
                if (sel && nombreInput) nombreInput.value = sel.dataset.nombre || '';
                if (sel && emailInput) emailInput.value = sel.dataset.email || '';

                // Reset dropzone
                selectedFiles = null;
                const dzText = document.querySelector('.dropzone .dz-text');
                if (dzText) dzText.textContent = 'Arrastre o';
                if (respInput) respInput.value = '';

                // Ensure checkboxes cleared
                const chkboxes = Array.from(document.querySelectorAll('.checkbox-row input[type="checkbox"]'));
                chkboxes.forEach(ch => ch.checked = false);
            }, totalDelay);
        } catch (err) {
            console.warn('animateClearPreserveUser error:', err);
        }
    }

    form.addEventListener("submit", async (e) => {
        e.preventDefault();

        const usuarioSeleccionado = usuarioSelect.options[usuarioSelect.selectedIndex];
        
        // Recolectar checkboxes seleccionadas del grupo
        const checked = Array.from(document.querySelectorAll('.checkbox-row input[type="checkbox"]:checked')).map(i => i.value);

        // If postgres mode active, prefer clone values if present
        const dbTypeValue = (document.getElementById('db_type') ? document.getElementById('db_type').value : 'oracle');
        const readCloneOrOrig = (id) => {
            if (dbTypeValue === 'postgres') {
                const clone = document.getElementById(id + '_pg_clone');
                if (clone) return clone.value.trim();
            }
            const orig = document.getElementById(id);
            return orig ? orig.value.trim() : '';
        };

        const checkmarksFinal = Array.from(document.querySelectorAll('.checkbox-row input[type="checkbox"]:checked')).map(i => i.value);
        // include postgres-specific checkbox if present
        const chkPg = document.getElementById('chk_postgres');
        if (chkPg && chkPg.checked) checkmarksFinal.push(chkPg.value);

        const data = {
            usuario: {
                id: usuarioSelect.value,
                nombre: usuarioSeleccionado.dataset.nombre,
                email: usuarioSeleccionado.dataset.email
            },
            solman: readCloneOrOrig('solman'),
            titulo: readCloneOrOrig('titulo'),
            ticket: readCloneOrOrig('ticket'),
            sistemas: readCloneOrOrig('sistemas'),
            checkmarks: checkmarksFinal, // Cambiado de checkmark a checkmarks
            procedure: readCloneOrOrig('procedure'),
            base_origen: (document.getElementById("base_origen") ? document.getElementById("base_origen").value : ''),
            esquema_origen: (document.getElementById("esquema_origen") ? document.getElementById("esquema_origen").value : ''),
            base_destino: (document.getElementById("base_destino") ? document.getElementById("base_destino").value : ''),
            esquema_destino: (document.getElementById("esquemas_destino") ? document.getElementById("esquemas_destino").value : ''),
            descripcion: document.getElementById("descripcion").value.trim(),
            resultado: document.getElementById("resultado").value.trim(),
            objetivo: document.getElementById("objetivo").value.trim(),
            // Use files selected through the dropzone if available, otherwise fall back to the input.files
            respaldos: (typeof selectedFiles !== 'undefined' && selectedFiles) ? selectedFiles : document.getElementById("respaldos").files
                ,
            db_type: dbTypeValue,
        };

        if (!data.usuario.id) {
            alert("Debes seleccionar un usuario.");
            return;
        }

        if (!data.solman) {
            alert("El número de Solman es obligatorio.");
            return;
        }

        // Verificar que las librerías necesarias estén cargadas
        if (typeof docx === 'undefined') {
            alert("Error: La librería para generar documentos no está disponible. Recarga la página.");
            return;
        }

        if (typeof JSZip === 'undefined') {
            alert("Error: La librería para generar ZIP no está disponible. Recarga la página.");
            return;
        }

        await generarZIP(data);

        // Animar la limpieza del formulario conservando el usuario
        animateClearPreserveUser();
    });
});
