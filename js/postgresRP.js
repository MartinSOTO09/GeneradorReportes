// Wrapper generator for Postgres Release Plan
// Delegates to the existing generarReleasePlan() implementation for now
// Kept as a separate file so you can add Postgres-specific generation later.
async function generarReleasePlan_postgres(data) {
    data = data || {};
    data.db_type = data.db_type || 'postgres';
    if (typeof generarReleasePlan === 'function') {
        return await generarReleasePlan(data);
    }
    throw new Error('generarReleasePlan is not defined. Ensure releasePlan.js is loaded before this file.');
}
