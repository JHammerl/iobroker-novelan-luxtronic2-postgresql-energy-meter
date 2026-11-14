/**
 * ============================================================================
 * Titel:        Wärmepumpen-Datenlogger (ioBroker -> PostgreSQL)
 * Projekt:      ioBroker Novelan/Luxtronic – PostgreSQL Energie- und Wärmemengenzähler
 * Version:      1.0
 * Beschreibung: Überwacht Strompreis, elektrische Leistungsaufnahme (kWh) und
 *               erzeugte Wärmemengen (kWh) der Wärmepumpe im RAM und schreibt
 *               die aktuellen Zählerstände über einen ausfallsicheren
 *               Verbindungspool (pg.Pool) zyklisch in PostgreSQL.
 *
 * Datenquelle:  ioBroker States (0_userdata.0.Wärmepumpe.* & Strompreis)
 * Ziel-DB:      PostgreSQL (DB: smarthome | Tabelle: heat_pump)
 *
 * Funktionsweise:
 *               - Sofort-Speicherung & Entwarnungs-Mail beim Skriptstart
 *               - Echtzeit-Aktualisierung des RAM-Puffers bei Wertänderung
 *               - Zyklischer Datenbank-Write alle 5 Minuten (300.000 ms)
 *               - Automatische Fehlerabfangung bei DB-Verbindungsabbruch
 *               - Sauberes Schließen des DB-Pools + Alarm-Mail bei Skript-Stopp
 *
 * NPM-Module:   pg (in den Adapter-Einstellungen eintragen)
 * ============================================================================
 */

const { Pool } = require('pg');

// ==================== KONFIGURATION ====================
// HIER DEINE EIGENEN ZUGANGSDATEN UND E-MAIL EINTRAGEN:
const pgConnectionString = 'postgresql://BENUTZERNAME:PASSWORT@192.168.1.100:5432/smarthome';
const measurement = 'heat_pump';
const SPEICHER_INTERVALL = 300000; // Alle 5 Minuten in Millisekunden
const ALERT_EMAIL = 'deine-email@beispiel.de'; // Empfänger für Start-/Stopp-Benachrichtigungen

// Pool einmalig global initialisieren (hält Verbindungen stabil und fängt Fehler ab)
const pool = new Pool({ 
    connectionString: pgConnectionString,
    max: 5, // Maximal 5 parallele Verbindungen reichen völlig aus
    idleTimeoutMillis: 30000, // Schließt ungenutzte Verbindungen nach 30 Sek
    connectionTimeoutMillis: 2000, // Bricht nach 2 Sek ab, falls DB nicht erreichbar
});

// Fehler-Handler für den Pool, damit das ioBroker-Skript bei DB-Ausfall nicht abstürzt
pool.on('error', (err) => {
    console.error('Unerwarteter PostgreSQL Pool-Fehler (heat_pump):', err.message);
});

// Zuordnung: ioBroker-Datenpunkt -> Spaltenname in PostgreSQL
const loggingTemplate = {
    '0_userdata.0.energy.contract.electricity.kWhPrice': 'priceIn',
    '0_userdata.0.Wärmepumpe.Leistungsaufnahme.Gesamt': 'GesamtLeistung-kWh',
    '0_userdata.0.Wärmepumpe.Leistungsaufnahme.Heizung': 'Heizung-kWh',
    '0_userdata.0.Wärmepumpe.Leistungsaufnahme.Warmwasser': 'Warmwasser-kWh',
    '0_userdata.0.Wärmepumpe.Wärmemenge.Gesamt': 'GesamteWärmenMengekWh',
    '0_userdata.0.Wärmepumpe.Wärmemenge.Heizung': 'Heizung-WM-kWh',
    '0_userdata.0.Wärmepumpe.Wärmemenge.Warmwasser': 'Warmwasser-WM-kWh',
};

const loggingObj = {};
let writeTimer = null;

// Hilfsfunktion für deutschen Zeitstempel
function getZeitpunkt() {
    return new Date().toLocaleDateString('de-DE', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit', second: '2-digit'
    });
}

// Zentrale Schreibfunktion
async function datenInDatabaseSchreiben() {
    const keys = Object.keys(loggingObj).filter(key => !isNaN(loggingObj[key]));
    if (keys.length === 0) return;

    const columns = keys.map(key => `"${key}"`).join(', ');
    const placeholders = keys.map((_, index) => `$${index + 1}`).join(', ');
    const values = keys.map(key => loggingObj[key]);

    const queryText = `INSERT INTO ${measurement} (${columns}) VALUES (${placeholders})`;

    try {
        await pool.query(queryText, values);
    } catch (err) {
        console.error("PostgreSQL Write Error (heat_pump):", err.message);
    }
}

async function start() {
    // 1. Initial alle Werte laden
    for (let [objId, key] of Object.entries(loggingTemplate)) {
        const state = await getStateAsync(objId);
        loggingObj[key] = (state && !isNaN(state.val)) ? state.val : 0;
    }

    // 2. Event-Listener: Aktualisiert NUR den Zwischenspeicher im RAM bei Änderung
    on({ id: Object.keys(loggingTemplate), change: 'ne' }, (obj) => {
        const key = loggingTemplate[obj.id];
        loggingObj[key] = obj.state.val;
    });

    // Sofort beim Skriptstart einmalig Daten schreiben
    await datenInDatabaseSchreiben();

    // Entwarnungs-Mail beim Starten absetzen
    if (ALERT_EMAIL && ALERT_EMAIL !== 'deine-email@beispiel.de') {
        sendTo("email.0", "send", {
            to: ALERT_EMAIL,
            subject: "✅ ENTWARNUNG: Wärmepumpen-Datenlogger (PostgreSQL) läuft wieder!",
            html: `
                <h3 style="color: #16a34a;">Entwarnung: Datenerfassung aktiv</h3>
                <p>Das ioBroker-Skript <b>Wärmepumpen-Datenlogger (heat_pump)</b> wurde am <b>${getZeitpunkt()} Uhr</b> erfolgreich gestartet.</p>
                <p>Die ersten aktuellen Zählerstände wurden bereits in die PostgreSQL-Datenbank geschrieben und die 5-Minuten-Aufzeichnung läuft wieder normal.</p>
            `
        });
    }
    log('Wärmepumpen-Datenlogger gestartet.');

    // 3. Intervall: Schreibt den aktuellen RAM-Zustand alle 5 Minuten stabil über den Pool
    writeTimer = setInterval(async () => {
        await datenInDatabaseSchreiben();
    }, SPEICHER_INTERVALL);
}

// 4. Sauberes Beenden + Alarm-Mail beim Stoppen oder Neustarten des Skripts
onStop(async (callback) => {
    if (writeTimer) {
        clearInterval(writeTimer);
    }

    if (ALERT_EMAIL && ALERT_EMAIL !== 'deine-email@beispiel.de') {
        sendTo("email.0", "send", {
            to: ALERT_EMAIL,
            subject: "⚠️ ALARM: Wärmepumpen-Datenlogger (PostgreSQL) wurde gestoppt!",
            html: `
                <h3 style="color: #dc2626;">Warnung: Datenerfassung unterbrochen</h3>
                <p>Das ioBroker-Skript <b>Wärmepumpen-Datenlogger (heat_pump)</b> wurde am <b>${getZeitpunkt()} Uhr</b> beendet oder neu gestartet.</p>
                <p>Bitte prüfe im ioBroker-Admin, ob das Skript wieder aktiv läuft, damit keine Lücken in der PostgreSQL-Datenbank entstehen.</p>
            `
        });
    }

    try {
        await pool.end();
        log('PostgreSQL Pool (heat_pump) sauber geschlossen.');
    } catch (err) {
        log('Fehler beim Schließen des PostgreSQL Pools: ' + err.message, 'error');
    }

    // Kurz 1 Sekunde warten, damit sendTo die Nachricht sicher an email.0 übergibt
    await new Promise(resolve => setTimeout(resolve, 1000));
    callback();
}, 3000);

start();
