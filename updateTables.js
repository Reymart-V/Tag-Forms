const sql = require('mssql');
const fs = require('fs');

const mills = [
    'ALD','BEM','BWK','CHA','COR','GP','GUN',
    'HL','HUG','JEF','JOA','LAS','NAC'
];

const HTML_PATH = '\\\\roc-fs01\\Global\\Reymart\\TagRequestForm.html';
const JSON_PATH = '\\\\roc-fs01\\Global\\Reymart\\tables.json';

const DB_OVERRIDES = {
    'JOA': 'JoannaDB',
    'JEF': 'JefProcessDB',
    'NAC': 'NacProcessDB',
    'COR': ['CordeleLine2', 'CordelePress1SQL']
};

function getConfig(server, database) {
    return {
        server: server,
        database: database,
        user: 'Reyadmin',
        password: 'Treymar421!',
        options: {
            trustServerCertificate: true,
            encrypt: false,
            enableArithAbort: true
        }
    };
}

async function getTablesFromDB(server, db) {
    const pool = await sql.connect(getConfig(server, db));
    const result = await pool.request().query(`
        SELECT TABLE_NAME 
        FROM INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_TYPE = 'BASE TABLE'
        ORDER BY TABLE_NAME
    `);
    await sql.close();
    return result.recordset.map(r => r.TABLE_NAME);
}

async function getTablesForMill(millCode) {
    const server = `${millCode}-HIST01`;
    const dbEntry = DB_OVERRIDES[millCode] || 'Historian';

    try {
        if (Array.isArray(dbEntry)) {
            console.log(`Connecting to ${server} (${dbEntry.join(', ')})...`);
            let allTables = [];
            for (const db of dbEntry) {
                try {
                    const tables = await getTablesFromDB(server, db);
                    console.log(`  Got ${tables.length} tables from ${db}`);
                    allTables = [...new Set([...allTables, ...tables])];
                } catch(e) {
                    try { await sql.close(); } catch(e2) {}
                    console.error(`  Failed on ${db}: ${e.message}`);
                }
            }
            return allTables;
        }

        console.log(`Connecting to ${server} (${dbEntry})...`);
        const tables = await getTablesFromDB(server, dbEntry);
        console.log(`  Got ${tables.length} tables`);
        return tables;

    } catch (err) {
        try { await sql.close(); } catch(e) {}
        console.error(`  Could not connect to ${server}: ${err.message}`);
        return [];
    }
}

async function main() {
    const allTables = {};
    for (const mill of mills) {
        allTables[mill] = await getTablesForMill(mill);
    }

    // Save JSON backup
    fs.writeFileSync(JSON_PATH, JSON.stringify(allTables, null, 2));
    console.log(`\nJSON saved to ${JSON_PATH}`);

    // Inject data directly into HTML file
    let html = fs.readFileSync(HTML_PATH, 'utf8');
    const placeholder = /let MILL_TABLES = \{\}; \/\/ MILL_TABLES_PLACEHOLDER/;
    const injection = `let MILL_TABLES = ${JSON.stringify(allTables)}; // MILL_TABLES_PLACEHOLDER`;

    if (placeholder.test(html)) {
        html = html.replace(placeholder, injection);
        fs.writeFileSync(HTML_PATH, html, 'utf8');
        console.log(`HTML updated at ${HTML_PATH}`);
    } else {
        console.error('Could not find placeholder in HTML file — make sure the HTML is the correct version');
    }

    console.log('\nDone!');
}

main();