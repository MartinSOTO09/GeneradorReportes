// Wrapper generator for Linux Test Evidence
// Delegates to the existing generarTestEvidence() implementation for now
// Kept as a separate file so you can add Linux-specific generation later.
async function generarTestEvidence_linux(data) {
    data = data || {};
    data.db_type = data.db_type || 'linux';
    if (typeof generarTestEvidence === 'function') {
        return await generarTestEvidence(data);
    }
    throw new Error('generarTestEvidence is not defined. Ensure TestEvidence.js is loaded before this file.');
}
