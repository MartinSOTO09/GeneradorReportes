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

        const data = {
            usuario: {
                id: usuarioSelect.value,
                nombre: usuarioSeleccionado.dataset.nombre,
                email: usuarioSeleccionado.dataset.email
            },
            solman: document.getElementById("solman").value.trim(),
            titulo: document.getElementById("titulo").value.trim(),
            ticket: document.getElementById("ticket").value.trim(),
            sistemas: document.getElementById("sistemas").value.trim(),
            checkmarks: checked, // Cambiado de checkmark a checkmarks
            procedure: document.getElementById("procedure").value.trim(),
            base_origen: document.getElementById("base_origen").value,
            esquema_origen: document.getElementById("esquema_origen").value,
            base_destino: document.getElementById("base_destino").value,
            esquema_destino: document.getElementById("esquemas_destino").value,
            descripcion: document.getElementById("descripcion").value.trim(),
            resultado: document.getElementById("resultado").value.trim(),
            // Use files selected through the dropzone if available, otherwise fall back to the input.files
            respaldos: (typeof selectedFiles !== 'undefined' && selectedFiles) ? selectedFiles : document.getElementById("respaldos").files
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
