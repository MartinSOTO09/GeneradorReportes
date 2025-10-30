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

    // Animación de la tarjeta de Información General al cambiar valores relevantes
    function triggerInfoCardAnim() {
        try {
            const infoCard = document.querySelector('section.card:nth-of-type(2)');
            if (!infoCard) return;
            infoCard.classList.add('card-change');
            setTimeout(() => infoCard.classList.remove('card-change'), 420);
        } catch (_) {}
    }

    // Escuchar cambios en campos clave de Información General
    const infoSelectors = ['#solman', '#ticket', '#titulo', '#base_origen', '#esquema_origen', '#base_destino', '#esquemas_destino', '#procedure', '#sistemas'];
    infoSelectors.forEach(sel => {
        const el = document.querySelector(sel);
        if (el) {
            el.addEventListener('change', triggerInfoCardAnim);
            el.addEventListener('input', () => {
                // limitar frecuencia de animación en input continuo
                if (!el.__animPending) {
                    el.__animPending = true;
                    triggerInfoCardAnim();
                    setTimeout(() => el.__animPending = false, 600);
                }
            });
        }
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
            // Añade una clase al body para facilitar estilos condicionales si se desea
            const applyDbBodyClass = (val) => {
                document.body.classList.remove('db-oracle', 'db-postgres', 'db-linux');
                document.body.classList.add(`db-${val}`);
            };

            const applyFieldsForDb = (val) => {
                const hideForExternal = (val === 'postgres' || val === 'linux');

                // Ocultar/mostrar contenedores de Base/Esquema y la fila de checkmarks
                ['#base_origen', '#esquema_origen', '#base_destino', '#esquemas_destino'].forEach(sel => {
                    const el = document.querySelector(sel);
                    if (!el) return;
                    const container = el.closest('div') || el;
                    container.style.display = hideForExternal ? 'none' : '';

                    // Evitar errores de validación HTML5 en campos ocultos
                    const isEsquema = (el.id === 'esquema_origen' || el.id === 'esquemas_destino');
                    if (hideForExternal) {
                        // Deshabilitar y quitar required si está presente
                        el.disabled = true;
                        if (isEsquema) el.required = false;
                    } else {
                        // Rehabilitar y restaurar required solo en esquemas
                        el.disabled = false;
                        if (isEsquema) el.required = true;
                    }
                });

                // Mostrar/Ocultar la sección completa de Check Mark según DB
                const chkRow = document.querySelector('.checkbox-row');
                if (chkRow) {
                    const chkContainer = chkRow.parentElement || chkRow;
                    chkContainer.style.display = hideForExternal ? 'none' : '';
                    // Si se muestra, asegurar el layout flex del row
                    if (!hideForExternal) chkRow.style.display = 'flex';
                }

                // Ajustar campos por tipo: procedure (Oracle), link (Postgres)
                const proc = document.getElementById('procedure');
                const linkInput = document.getElementById('link');
                if (proc) {
                    try {
                        // Cambiar etiqueta asociada al campo según DB
                        const procLabel = proc.parentNode && proc.parentNode.querySelector && proc.parentNode.querySelector('label');
                        const procWrapper = proc.parentElement || proc;
                        const linkWrapper = linkInput ? (linkInput.parentElement || linkInput) : null;
                        const linkLabel = linkWrapper && linkWrapper.querySelector ? linkWrapper.querySelector('label') : null;
                        // Containers para reordenar en fila 2
                        const solicitudCont = document.getElementById('solicitud_container');
                        const nombreSolicitudCont = document.getElementById('nombre_solicitud_container');
                        const sistemasEl = document.getElementById('sistemas');
                        const sistemasWrapper = sistemasEl ? (sistemasEl.parentElement || sistemasEl) : null;
                        const row2 = document.querySelector('.row2');
                        // Postgres: usar campo Link dedicado y ocultar Procedure
                        if (val === 'postgres') {
                            // Ocultar y des-requerir Procedure
                            proc.required = false;
                            if (procWrapper) {
                                procWrapper.style.display = 'none';
                                procWrapper.style.removeProperty('order');
                            }
                            // Configurar Link
                            if (linkInput) {
                                linkInput.type = 'url';
                                linkInput.placeholder = 'Pega aquí el enlace (ej: https://... o ssh://...)';
                                linkInput.style.width = '100%';
                                linkInput.style.maxWidth = '1000px';
                                linkInput.required = true;
                            }
                            if (linkWrapper) {
                                linkWrapper.style.display = '';
                                linkWrapper.style.removeProperty('grid-column');
                                linkWrapper.style.order = '3';
                            }
                            if (linkLabel) linkLabel.textContent = 'Link';

                            // Mostrar Solicitud y reordenar: Solicitud, Sistema, Link (Link más ancho)
                            if (solicitudCont) solicitudCont.style.display = '';
                            if (row2) row2.style.gridTemplateColumns = '200px 200px 1fr';
                            if (solicitudCont) solicitudCont.style.order = '1';
                            if (sistemasWrapper) sistemasWrapper.style.order = '2';
                            // Ocultar el campo Nombre Solicitud en Postgres
                            if (nombreSolicitudCont) {
                                nombreSolicitudCont.style.display = 'none';
                                nombreSolicitudCont.style.removeProperty('order');
                            }
                        // Linux: solo mostrar Solicitud y Sistema; ocultar el campo de procedure
                        } else if (val === 'linux') {
                            proc.type = 'text';
                            proc.placeholder = 'Solicitud(es)';
                            proc.style.width = '100%';
                            proc.style.maxWidth = '1000px';
                            if (procWrapper) procWrapper.style.removeProperty('grid-column');
                            if (procLabel) procLabel.textContent = 'Solicitud(es)';
                            // En Linux ocultamos el campo procedure; debe NO ser requerido para no bloquear la validación nativa
                            proc.required = false;
                            // Ocultar Link en Linux
                            if (linkWrapper) {
                                linkWrapper.style.display = 'none';
                                linkWrapper.style.removeProperty('order');
                            }
                            if (linkInput) linkInput.required = false;

                            // Mostrar Solicitud y Nombre Solicitud; ocultar procedure; configurar tres columnas
                            if (solicitudCont) solicitudCont.style.display = '';
                            if (nombreSolicitudCont) nombreSolicitudCont.style.display = '';
                            if (row2) row2.style.gridTemplateColumns = '200px 200px 200px';
                            if (solicitudCont) solicitudCont.style.order = '1';
                            if (nombreSolicitudCont) nombreSolicitudCont.style.order = '2';
                            if (sistemasWrapper) sistemasWrapper.style.order = '3';
                            if (procWrapper) {
                                procWrapper.style.order = '3';
                                procWrapper.style.display = 'none';
                            }
                        // Oracle: valor por defecto
                        } else {
                            proc.type = 'text';
                            proc.placeholder = 'Nombre del procedimiento';
                            // Restaurar ancho por defecto
                            proc.style.removeProperty('max-width');
                            proc.style.removeProperty('width');
                            if (procWrapper) procWrapper.style.removeProperty('grid-column');
                            if (procLabel) procLabel.textContent = 'Procedure';
                            // En Oracle el campo procedure vuelve a ser requerido
                            proc.required = true;
                            // Ocultar Link en Oracle
                            if (linkWrapper) {
                                linkWrapper.style.display = 'none';
                                linkWrapper.style.removeProperty('order');
                            }
                            if (linkInput) linkInput.required = false;

                            // Ocultar Solicitud y restaurar orden/plantilla de columnas
                            const solicitudCont2 = document.getElementById('solicitud_container');
                            if (solicitudCont2) solicitudCont2.style.display = 'none';
                            if (nombreSolicitudCont) {
                                nombreSolicitudCont.style.display = 'none';
                                nombreSolicitudCont.style.removeProperty('order');
                            }
                            if (row2) row2.style.gridTemplateColumns = '200px 200px 1fr 200px';
                            // Reset de order para evitar efectos secundarios
                            if (solicitudCont2) solicitudCont2.style.removeProperty('order');
                            if (sistemasWrapper) sistemasWrapper.style.removeProperty('order');
                            if (procWrapper) {
                                procWrapper.style.removeProperty('order');
                                // Asegurar que el campo procedure se muestre en Oracle
                                procWrapper.style.display = '';
                            }
                        }
                    } catch (_) { /* noop */ }
                }
            };

            // Inicializar según valor actual y registrar listener
            const initial = dbTypeSelect.value || 'oracle';
            applyDbBodyClass(initial);
            applyFieldsForDb(initial);
            // pequeña animación inicial
            dbTypeSelect.classList.add('pulse');
            setTimeout(() => dbTypeSelect.classList.remove('pulse'), 220);

            const dbIcon = document.getElementById('dbIcon');
            dbTypeSelect.addEventListener('change', (e) => {
                const v = e.target.value || 'oracle';
                applyDbBodyClass(v);
                applyFieldsForDb(v);
                // animación sutil en el select
                dbTypeSelect.classList.add('pulse');
                setTimeout(() => dbTypeSelect.classList.remove('pulse'), 220);
                // animación sutil en el icono
                if (dbIcon) {
                    dbIcon.style.transform = 'scale(1.12)';
                    setTimeout(() => { dbIcon.style.transform = 'scale(1)'; }, 160);
                }
                triggerInfoCardAnim();
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

    // Botón Limpiar: limpiar solo cuando el usuario lo decida
    const btnClear = document.getElementById('btn_clear');
    if (btnClear) {
        btnClear.addEventListener('click', (e) => {
            e.preventDefault();
            animateClearPreserveUser();
        });
    }

    form.addEventListener("submit", async (e) => {
        e.preventDefault();

        const usuarioSeleccionado = usuarioSelect.options[usuarioSelect.selectedIndex];
        
        // Recolectar checkboxes seleccionadas del grupo
        const checked = Array.from(document.querySelectorAll('.checkbox-row input[type="checkbox"]:checked')).map(i => i.value);

        const dbTypeValue = (document.getElementById('db_type') ? document.getElementById('db_type').value : 'oracle');

        const checkmarksFinal = Array.from(document.querySelectorAll('.checkbox-row input[type="checkbox"]:checked')).map(i => i.value);

        const data = {
            usuario: {
                id: usuarioSelect.value,
                nombre: usuarioSeleccionado.dataset.nombre,
                email: usuarioSeleccionado.dataset.email
            },
            solman: (document.getElementById('solman') ? document.getElementById('solman').value.trim() : ''),
            titulo: (document.getElementById('titulo') ? document.getElementById('titulo').value.trim() : ''),
            ticket: (document.getElementById('ticket') ? document.getElementById('ticket').value.trim() : ''),
            sistemas: (document.getElementById('sistemas') ? document.getElementById('sistemas').value.trim() : ''),
            checkmarks: checkmarksFinal, // Cambiado de checkmark a checkmarks
            procedure: (document.getElementById('procedure') ? document.getElementById('procedure').value.trim() : ''),
            base_origen: (document.getElementById("base_origen") ? document.getElementById("base_origen").value : ''),
            esquema_origen: (document.getElementById("esquema_origen") ? document.getElementById("esquema_origen").value : ''),
            base_destino: (document.getElementById("base_destino") ? document.getElementById("base_destino").value : ''),
            esquema_destino: (document.getElementById("esquemas_destino") ? document.getElementById("esquemas_destino").value : ''),
            descripcion: document.getElementById("descripcion").value.trim(),
            resultado: document.getElementById("resultado").value.trim(),
            objetivo: document.getElementById("objetivo").value.trim(),
            solicitud: (document.getElementById('solicitud') ? document.getElementById('solicitud').value.trim() : ''),
            nombre_solicitud: (document.getElementById('nombre_solicitud') ? document.getElementById('nombre_solicitud').value.trim() : ''),
            link: (document.getElementById('link') ? document.getElementById('link').value.trim() : ''),
            // Use files selected through the dropzone if available, otherwise fall back to the input.files
            respaldos: (typeof selectedFiles !== 'undefined' && selectedFiles) ? selectedFiles : document.getElementById("respaldos").files
                ,
            db_type: dbTypeValue,
        };

        // Compatibilidad: en Postgres, mapear link -> procedure para los generadores existentes
        if (data.db_type === 'postgres') {
            data.procedure = data.link;
        }

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
    });

});
