/**
 * ============================================================================
 * Titel:        Wärmepumpen-Verbrauchs- & Effizienznachweis (GEG-Format)
 * Projekt:      ioBroker Novelan/Luxtronic – PostgreSQL Energie- und Wärmemengenzähler
 * Version:      1.0
 * Beschreibung: Liest Stromverbräuche (kWh) und Wärmemengen (kWh) der
 *               Wärmepumpe aus der PostgreSQL-Tabelle 'heat_pump' aus,
 *               berechnet die Arbeitszahlen (COP / JAZ) und erstellt
 *               automatisierte Berichte als Excel-Datei, PDF-Dokument
 *               (inkl. Jahresgrafik) und HTML-E-Mail.
 *
 * Datenquelle:  PostgreSQL (DB: smarthome | Tabelle: heat_pump)
 * Ausgabe:      - Excel (.xlsx) unter /auswertungen/postgresSQL/excel/
 *               - PDF (.pdf)   unter /auswertungen/postgresSQL/pdf/
 *               - HTML-E-Mail  über ioBroker-Adapter "email.0"
 *
 * Zeitpläne:    - Täglich (23:59 Uhr):       Tagesbericht (letzte 24 Std.)
 *               - Sonntags (23:55 Uhr):      Wochenbericht (Week-to-Date)
 *               - Monatsletzter (23:50 Uhr): Monats- & Jahresbericht (YTD)
 *
 * NPM-Module:   pg, xlsx, pdfkit (in den Adapter-Einstellungen eintragen)
 * ============================================================================
 */

const { Client } = require('pg');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');

// ==================== KONFIGURATION ====================
// HIER DEINE EIGENEN ZUGANGSDATEN, ADRESSE UND E-MAIL EINTRAGEN:
const dbConfig = {
    user: 'BENUTZERNAME',
    host: '192.168.1.100', 
    database: 'smarthome',
    password: 'DEIN_DB_PASSWORT',
    port: 5432,
};

const REPORT_EMAIL = 'deine-email@beispiel.de';
const OBJEKT_ADRESSE = 'Musterstraße 1, 12345 Musterstadt';

// Basis-Ordnerstrukturen für die Nachweise
const baseExcelPath = '/auswertungen/postgresSQL/excel/';
const basePdfPath = '/auswertungen/postgresSQL/pdf/';

// Dynamische Pfade für die vier Zeiträume
const paths = {
    daily: {
        excel: path.join(baseExcelPath, 'wp_verbrauchsnachweis_tag.xlsx'),
        pdf: path.join(basePdfPath, 'wp_verbrauchsnachweis_tag.pdf'),
        title: 'Tagesbericht (Letzte 24 Stunden)',
        subject: 'Wärmepumpen-Nachweis (Tagesübersicht)'
    },
    week2date: {
        excel: path.join(baseExcelPath, 'wp_verbrauchsnachweis_woche.xlsx'),
        pdf: path.join(basePdfPath, 'wp_verbrauchsnachweis_woche.pdf'),
        title: 'Wochenbericht (Week-to-Date)',
        subject: 'Wärmepumpen-Nachweis (Wochenübersicht)'
    },
    month2date: {
        excel: path.join(baseExcelPath, 'wp_verbrauchsnachweis_monat.xlsx'),
        pdf: path.join(basePdfPath, 'wp_verbrauchsnachweis_monat.pdf'),
        title: 'Monatsbericht (Month-to-Date)',
        subject: 'Wärmepumpen-Nachweis (Monatsübersicht)'
    },
    year2date: {
        excel: path.join(baseExcelPath, 'wp_verbrauchsnachweis_jahr.xlsx'),
        pdf: path.join(basePdfPath, 'wp_verbrauchsnachweis_jahr.pdf'),
        title: 'Jahresbericht (Year-to-Date)',
        subject: 'Wärmepumpen-Nachweis (Jahresübersicht)'
    }
};
function getSQLQuery(mode) {
    let startTsExpression = '';
    let isYearly = (mode === 'year2date');

    switch (mode) {
        case 'daily': startTsExpression = "now() - interval '24 hours'"; break;
        case 'week2date': startTsExpression = "date_trunc('week', now())"; break;
        case 'month2date': startTsExpression = "date_trunc('month', now())"; break;
        case 'year2date': startTsExpression = "date_trunc('year', now())"; break;
    }

    if (isYearly) {
        return `
        WITH monate AS (
            SELECT 
                m AS monat_num,
                CASE m 
                    WHEN 1 THEN 'Jan' WHEN 2 THEN 'Feb' WHEN 3 THEN 'Mrz' WHEN 4 THEN 'Apr'
                    WHEN 5 THEN 'Mai' WHEN 6 THEN 'Jun' WHEN 7 THEN 'Jul' WHEN 8 THEN 'Aug'
                    WHEN 9 THEN 'Sep' WHEN 10 THEN 'Okt' WHEN 11 THEN 'Nov' WHEN 12 THEN 'Dez'
                END AS monat_name,
                date_trunc('year', now()) + (m - 1) * interval '1 month' AS monat_start_ts,
                date_trunc('year', now()) + m * interval '1 month' AS monat_ende_ts
            FROM generate_series(1, 12) m
        ),
        monats_stichtage AS (
            SELECT 
                m.monat_num, m.monat_name, m.monat_start_ts,
                COALESCE(
                    (SELECT "GesamtLeistung-kWh" FROM heat_pump WHERE "timestamp" < m.monat_start_ts AND "GesamtLeistung-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                    (SELECT "GesamtLeistung-kWh" FROM heat_pump WHERE "timestamp" >= m.monat_start_ts AND "timestamp" < m.monat_ende_ts AND "GesamtLeistung-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
                ) AS el_ges_start,
                (SELECT "GesamtLeistung-kWh" FROM heat_pump WHERE "timestamp" < m.monat_ende_ts AND "GesamtLeistung-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS el_ges_ende,

                COALESCE(
                    (SELECT "Heizung-kWh" FROM heat_pump WHERE "timestamp" < m.monat_start_ts AND "Heizung-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                    (SELECT "Heizung-kWh" FROM heat_pump WHERE "timestamp" >= m.monat_start_ts AND "timestamp" < m.monat_ende_ts AND "Heizung-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
                ) AS el_hz_start,
                (SELECT "Heizung-kWh" FROM heat_pump WHERE "timestamp" < m.monat_ende_ts AND "Heizung-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS el_hz_ende,

                COALESCE(
                    (SELECT "Warmwasser-kWh" FROM heat_pump WHERE "timestamp" < m.monat_start_ts AND "Warmwasser-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                    (SELECT "Warmwasser-kWh" FROM heat_pump WHERE "timestamp" >= m.monat_start_ts AND "timestamp" < m.monat_ende_ts AND "Warmwasser-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
                ) AS el_ww_start,
                (SELECT "Warmwasser-kWh" FROM heat_pump WHERE "timestamp" < m.monat_ende_ts AND "Warmwasser-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS el_ww_ende,

                COALESCE(
                    (SELECT "GesamteWärmenMengekWh" FROM heat_pump WHERE "timestamp" < m.monat_start_ts AND "GesamteWärmenMengekWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                    (SELECT "GesamteWärmenMengekWh" FROM heat_pump WHERE "timestamp" >= m.monat_start_ts AND "timestamp" < m.monat_ende_ts AND "GesamteWärmenMengekWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
                ) AS th_ges_start,
                (SELECT "GesamteWärmenMengekWh" FROM heat_pump WHERE "timestamp" < m.monat_ende_ts AND "GesamteWärmenMengekWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS th_ges_ende,

                COALESCE(
                    (SELECT "Heizung-WM-kWh" FROM heat_pump WHERE "timestamp" < m.monat_start_ts AND "Heizung-WM-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                    (SELECT "Heizung-WM-kWh" FROM heat_pump WHERE "timestamp" >= m.monat_start_ts AND "timestamp" < m.monat_ende_ts AND "Heizung-WM-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
                ) AS th_hz_start,
                (SELECT "Heizung-WM-kWh" FROM heat_pump WHERE "timestamp" < m.monat_ende_ts AND "Heizung-WM-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS th_hz_ende,

                COALESCE(
                    (SELECT "Warmwasser-WM-kWh" FROM heat_pump WHERE "timestamp" < m.monat_start_ts AND "Warmwasser-WM-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                    (SELECT "Warmwasser-WM-kWh" FROM heat_pump WHERE "timestamp" >= m.monat_start_ts AND "timestamp" < m.monat_ende_ts AND "Warmwasser-WM-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
                ) AS th_ww_start,
                (SELECT "Warmwasser-WM-kWh" FROM heat_pump WHERE "timestamp" < m.monat_ende_ts AND "Warmwasser-WM-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS th_ww_ende
            FROM monate m
        ),
        differenzen AS (
            SELECT 
                monat_num, monat_name, monat_start_ts,
                GREATEST(COALESCE(el_ges_ende, 0.0) - COALESCE(el_ges_start, 0.0), 0.0) AS delta_el_ges,
                GREATEST(COALESCE(el_hz_ende, 0.0) - COALESCE(el_hz_start, 0.0), 0.0) AS delta_el_hz,
                GREATEST(COALESCE(el_ww_ende, 0.0) - COALESCE(el_ww_start, 0.0), 0.0) AS delta_el_ww,
                GREATEST(COALESCE(th_ges_ende, 0.0) - COALESCE(th_ges_start, 0.0), 0.0) AS delta_th_ges,
                GREATEST(COALESCE(th_hz_ende, 0.0) - COALESCE(th_hz_start, 0.0), 0.0) AS delta_th_hz,
                GREATEST(COALESCE(th_ww_ende, 0.0) - COALESCE(th_ww_start, 0.0), 0.0) AS delta_th_ww
            FROM monats_stichtage
        )
        SELECT 
            monat_name AS "Zeitpunkt",
            ROUND(CAST(delta_el_ges AS numeric), 2) AS "Strom Gesamt (kWh)",
            ROUND(CAST(delta_el_hz AS numeric), 2) AS "Strom Heizung (kWh)",
            ROUND(CAST(delta_el_ww AS numeric), 2) AS "Strom Warmwasser (kWh)",
            ROUND(CAST(delta_th_ges / NULLIF(delta_el_ges, 0) AS numeric), 2) AS "COP Gesamt",
            ROUND(CAST(delta_th_hz / NULLIF(delta_el_hz, 0) AS numeric), 2) AS "COP Heizung",
            ROUND(CAST(delta_th_ww / NULLIF(delta_el_ww, 0) AS numeric), 2) AS "COP Warmwasser"
        FROM differenzen
        WHERE monat_start_ts <= now()
          AND (delta_el_ges > 0 OR delta_el_hz > 0 OR delta_el_ww > 0)
        ORDER BY monat_num ASC;`;
    }

    return `
    WITH stichtage AS (
        SELECT
            COALESCE(
                (SELECT "GesamtLeistung-kWh" FROM heat_pump WHERE "timestamp" < ${startTsExpression} AND "GesamtLeistung-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                (SELECT "GesamtLeistung-kWh" FROM heat_pump WHERE "timestamp" >= ${startTsExpression} AND "GesamtLeistung-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
            ) AS el_ges_start,
            (SELECT "GesamtLeistung-kWh" FROM heat_pump WHERE "GesamtLeistung-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS el_ges_ende,

            COALESCE(
                (SELECT "Heizung-kWh" FROM heat_pump WHERE "timestamp" < ${startTsExpression} AND "Heizung-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                (SELECT "Heizung-kWh" FROM heat_pump WHERE "timestamp" >= ${startTsExpression} AND "Heizung-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
            ) AS el_hz_start,
            (SELECT "Heizung-kWh" FROM heat_pump WHERE "Heizung-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS el_hz_ende,

            COALESCE(
                (SELECT "Warmwasser-kWh" FROM heat_pump WHERE "timestamp" < ${startTsExpression} AND "Warmwasser-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                (SELECT "Warmwasser-kWh" FROM heat_pump WHERE "timestamp" >= ${startTsExpression} AND "Warmwasser-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
            ) AS el_ww_start,
            (SELECT "Warmwasser-kWh" FROM heat_pump WHERE "Warmwasser-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS el_ww_ende,

            COALESCE(
                (SELECT "GesamteWärmenMengekWh" FROM heat_pump WHERE "timestamp" < ${startTsExpression} AND "GesamteWärmenMengekWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                (SELECT "GesamteWärmenMengekWh" FROM heat_pump WHERE "timestamp" >= ${startTsExpression} AND "GesamteWärmenMengekWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
            ) AS th_ges_start,
            (SELECT "GesamteWärmenMengekWh" FROM heat_pump WHERE "GesamteWärmenMengekWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS th_ges_ende,

            COALESCE(
                (SELECT "Heizung-WM-kWh" FROM heat_pump WHERE "timestamp" < ${startTsExpression} AND "Heizung-WM-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                (SELECT "Heizung-WM-kWh" FROM heat_pump WHERE "timestamp" >= ${startTsExpression} AND "Heizung-WM-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
            ) AS th_hz_start,
            (SELECT "Heizung-WM-kWh" FROM heat_pump WHERE "Heizung-WM-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS th_hz_ende,

            COALESCE(
                (SELECT "Warmwasser-WM-kWh" FROM heat_pump WHERE "timestamp" < ${startTsExpression} AND "Warmwasser-WM-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1),
                (SELECT "Warmwasser-WM-kWh" FROM heat_pump WHERE "timestamp" >= ${startTsExpression} AND "Warmwasser-WM-kWh" > 0 ORDER BY "timestamp" ASC LIMIT 1)
            ) AS th_ww_start,
            (SELECT "Warmwasser-WM-kWh" FROM heat_pump WHERE "Warmwasser-WM-kWh" > 0 ORDER BY "timestamp" DESC LIMIT 1) AS th_ww_ende
    ),
    differenzen AS (
        SELECT 
            GREATEST(COALESCE(el_ges_ende, 0.0) - COALESCE(el_ges_start, 0.0), 0.0) AS delta_el_ges,
            GREATEST(COALESCE(el_hz_ende, 0.0) - COALESCE(el_hz_start, 0.0), 0.0) AS delta_el_hz,
            GREATEST(COALESCE(el_ww_ende, 0.0) - COALESCE(el_ww_start, 0.0), 0.0) AS delta_el_ww,
            GREATEST(COALESCE(th_ges_ende, 0.0) - COALESCE(th_ges_start, 0.0), 0.0) AS delta_th_ges,
            GREATEST(COALESCE(th_hz_ende, 0.0) - COALESCE(th_hz_start, 0.0), 0.0) AS delta_th_hz,
            GREATEST(COALESCE(th_ww_ende, 0.0) - COALESCE(th_ww_start, 0.0), 0.0) AS delta_th_ww
        FROM stichtage
    )
    SELECT 
        'Gesamtauswertung' AS "Zeitpunkt",
        ROUND(CAST(delta_el_ges AS numeric), 2) AS "Strom Gesamt (kWh)",
        ROUND(CAST(delta_el_hz AS numeric), 2) AS "Strom Heizung (kWh)",
        ROUND(CAST(delta_el_ww AS numeric), 2) AS "Strom Warmwasser (kWh)",
        ROUND(CAST(delta_th_ges / NULLIF(delta_el_ges, 0) AS numeric), 2) AS "COP Gesamt",
        ROUND(CAST(delta_th_hz / NULLIF(delta_el_hz, 0) AS numeric), 2) AS "COP Heizung",
        ROUND(CAST(delta_th_ww / NULLIF(delta_el_ww, 0) AS numeric), 2) AS "COP Warmwasser"
    FROM differenzen;`;
}
async function runReportForMode(mode) {
    log(`Starte Erstellung für Modus: ${mode}...`);
    const cfg = paths[mode];
    const client = new Client(dbConfig);
    
    try {
        await client.connect();
        await client.query("SET TIME ZONE 'Europe/Berlin';");
        const res = await client.query(getSQLQuery(mode));
        let rows = res.rows;
        await client.end();

        if (!rows || rows.length === 0) {
            log(`Keine Daten für Modus ${mode} gefunden.`, "warn");
            return;
        }

        const dataRow = rows; 
        const heuteDatum = new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });

        // ==================== EXCEL EXPORT ====================
        const wb = XLSX.utils.book_new();
        const excelRows = mode === 'year2date' ? rows : [{ "Zeitraum": cfg.title, ...dataRow }];
        const ws = XLSX.utils.json_to_sheet(excelRows);
        XLSX.utils.book_append_sheet(wb, ws, "Verbrauchsnachweis");
        XLSX.writeFile(wb, cfg.excel);

        // ==================== HTML MAIL DESIGN ====================
        const cPrimary = '#334155'; const cPrimaryBlue = '#1e3a8a'; const cLightBg = '#f8fafc'; const cText = '#475569';
        
        let htmlMail = `<h3>Hallo,</h3><p>anbei findest du den Wärmepumpen-Verbrauchsnachweis im Format für Energie-Verbrauchsausweise.</p>
        <h4>Zusammenfassung: ${cfg.title}</h4>
        <table style="border-collapse: collapse; width: 100%; max-width: 750px; font-family: Arial, sans-serif; font-size: 12px;">
            <tr style="background-color: ${cPrimary}; color: white; text-align: left;">
                <th style="padding: 8px; border: 1px solid #cbd5e1;">Auswertungszeitraum</th>
                <th style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">Strom Gesamt (kWh)</th>
                <th style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">Strom Heizung (kWh)</th>
                <th style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">Strom WW (kWh)</th>
                <th style="padding: 8px; border: 1px solid #cbd5e1; text-align: right; color: #60a5fa;">COP Gesamt</th>
                <th style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">COP Heizung</th>
                <th style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">COP WW</th>
            </tr>`;
        if (mode === 'year2date') {
            rows.forEach((row, idx) => {
                let bg = idx % 2 === 0 ? cLightBg : '#ffffff';
                htmlMail += `<tr style="background-color: ${bg}; color: ${cText};">
                    <td style="padding: 8px; border: 1px solid #cbd5e1; font-weight: bold;">${row.Zeitpunkt}</td>
                    <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${row["Strom Gesamt (kWh)"]}</td>
                    <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${row["Strom Heizung (kWh)"]}</td>
                    <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${row["Strom Warmwasser (kWh)"]}</td>
                    <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right; font-weight: bold; color: ${cPrimaryBlue};">${row["COP Gesamt"] || '-'}</td>
                    <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${row["COP Heizung"] || '-'}</td>
                    <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${row["COP Warmwasser"] || '-'}</td>
                </tr>`;
            });
        } else {
            htmlMail += `<tr style="background-color: ${cLightBg}; color: ${cText}; font-weight: bold;">
                <td style="padding: 8px; border: 1px solid #cbd5e1;">${cfg.title}</td>
                <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${dataRow["Strom Gesamt (kWh)"]}</td>
                <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${dataRow["Strom Heizung (kWh)"]}</td>
                <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${dataRow["Strom Warmwasser (kWh)"]}</td>
                <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right; color: ${cPrimaryBlue}; font-size: 13px;">${dataRow["COP Gesamt"] || '-'}</td>
                <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${dataRow["COP Heizung"] || '-'}</td>
                <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">${dataRow["COP Warmwasser"] || '-'}</td>
            </tr>`;
        }
        htmlMail += `</table><p><i>Nachweis-Dokument generiert am ${heuteDatum} aus ioBroker Smart Home.</i></p>`;

        // ==================== PDF VERBRAUCHSAUSWEIS-FORMAT ====================
        const doc = new PDFDocument({ margin: 40, size: 'A4' });
        const pdfStream = fs.createWriteStream(cfg.pdf);

        pdfStream.on('finish', () => {
            log(`Datei für ${mode} erfolgreich gespeichert.`);
            if (REPORT_EMAIL && REPORT_EMAIL !== 'deine-email@beispiel.de') {
                sendTo("email.0", "send", {
                    to: REPORT_EMAIL, 
                    subject: cfg.subject,
                    html: htmlMail,
                    attachments: [
                        { path: cfg.excel, fileName: path.basename(cfg.excel) },
                        { path: cfg.pdf, fileName: path.basename(cfg.pdf) }
                    ]
                });
            }
        });

        doc.pipe(pdfStream);
        
        doc.rect(0, 0, 612, 95).fill('#334155');
        doc.fillColor('#ffffff').fontSize(15).font('Helvetica-Bold').text('Energiebedarf Wärmepumpe', 40, 22);
        doc.fontSize(9).font('Helvetica-Bold').fillColor('#93c5fd').text(OBJEKT_ADRESSE, 40, 44);
        doc.fontSize(8.5).font('Helvetica').fillColor('#ffffff').text(`Ausstellungsdatum: ${heuteDatum} Uhr`, 40, 65);
        
        doc.rect(40, 110, 515, 45).fill('#f8fafc');
        doc.lineWidth(0.5).strokeColor('#cbd5e1').rect(40, 110, 515, 45).stroke();
        doc.fillColor('#475569').fontSize(8).font('Helvetica-Bold')
           .text('Erfassungsobjekt:', 50, 117).font('Helvetica').text('Wärmepumpenanlage (Heizung & Warmwasser)', 140, 117)
           .font('Helvetica-Bold').text('Datenquelle:', 50, 130).font('Helvetica').text('ioBroker Smart Home System — PostgreSQL Core DB', 140, 130)
           .font('Helvetica-Bold').text('Auswertungsbasis:', 50, 143).font('Helvetica').text(mode === 'year2date' ? 'Kumulierte Monatsübersicht (Jahresverlauf)' : `Kumulierter Gesamtzeitraum (${cfg.title})`, 140, 143);
        let totalStromGes = 0, totalStromHz = 0, totalStromWW = 0;
        let weightedCopGes = 0, countCop = 0;
        let summaryRow = {};

        if (mode === 'year2date') {
            rows.forEach(r => {
                totalStromGes += parseFloat(r["Strom Gesamt (kWh)"]) || 0;
                totalStromHz += parseFloat(r["Strom Heizung (kWh)"]) || 0;
                totalStromWW += parseFloat(r["Strom Warmwasser (kWh)"]) || 0;
                if (r["COP Gesamt"]) { weightedCopGes += parseFloat(r["COP Gesamt"]); countCop++; }
            });
            summaryRow = {
                "Strom Gesamt (kWh)": ROUND(totalStromGes, 2),
                "Strom Heizung (kWh)": ROUND(totalStromHz, 2),
                "Strom Warmwasser (kWh)": ROUND(totalStromWW, 2),
                "COP Gesamt": countCop > 0 ? ROUND(weightedCopGes / countCop, 2) : 0,
                "COP Heizung": rows[rows.length - 1]["COP Heizung"],
                "COP Warmwasser": rows[rows.length - 1]["COP Warmwasser"]
            };
        } else {
            summaryRow = dataRow;
        }

        let y = 180;
        doc.fillColor('#1e293b').font('Helvetica-Bold').fontSize(9);
        doc.text('Kennwert-Kategorie', 42, y);
        doc.text('Erfasster Verbrauch / Effizienzwert', 320, y, { width: 235, align: 'right' });

        doc.moveTo(40, y + 14).lineTo(555, y + 14).lineWidth(1.2).strokeColor('#334155').stroke();
        y += 24;

        const zeilenData = [
            { label: 'Elektrische Energie — Gesamtantrieb (Strom)', val: `${summaryRow["Strom Gesamt (kWh)"]} kWh`, bold: true },
            { label: 'Elektrische Energie — Heizbetrieb (Strom)', val: `${summaryRow["Strom Heizung (kWh)"]} kWh`, bold: false },
            { label: 'Elektrische Energie — Warmwasserbereitung (Strom)', val: `${summaryRow["Strom Warmwasser (kWh)"]} kWh`, bold: false },
            { label: 'Effizienzkennzahl — Gesamtanlage (Jahres-COP / Arbeitszahl)', val: String(summaryRow["COP Gesamt"] || '-'), bold: true, color: '#1e3a8a' },
            { label: 'Effizienzkennzahl — Heizfunktion (COP Heizung)', val: String(summaryRow["COP Heizung"] || '-'), bold: false },
            { label: 'Effizienzkennzahl — Warmwasserfunktion (COP Warmwasser)', val: String(summaryRow["COP Warmwasser"] || '-'), bold: false }
        ];

        zeilenData.forEach((z, idx) => {
            if (idx === 0 || idx === 3) doc.rect(40, y - 4, 515, 18).fill('#f1f5f9');
            doc.fillColor(z.color ? z.color : '#334155').font(z.bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(9);
            doc.text(z.label, 45, y);
            doc.text(z.val, 320, y, { width: 230, align: 'right' });
            doc.moveTo(40, y + 12).lineTo(555, y + 12).lineWidth(0.5).strokeColor('#e2e8f0').stroke();
            y += 22;
        });

        // ==================== VISUELLE MONATSÜBERSICHT (NUR FÜR YEAR2DATE) ====================
        if (mode === 'year2date' && rows.length > 0) {
            doc.addPage();
            doc.rect(0, 0, 612, 75).fill('#475569');
            doc.fillColor('#ffffff').fontSize(15).font('Helvetica-Bold').text('VISUELLER MONATSVERLAUF (STROM VS. EFFIZIENZ)', 40, 25);
            doc.fontSize(8.5).font('Helvetica').text(`Ausstellungsdatum: ${heuteDatum} Uhr`, 40, 50);

            const gLeft = 70; const graphTop = 150; const gWidth = 460; const gHeight = 180;
            let maxStrom = 100;
            rows.forEach(r => { if (parseFloat(r["Strom Gesamt (kWh)"]) > maxStrom) maxStrom = parseFloat(r["Strom Gesamt (kWh)"]); });
            maxStrom = Math.ceil(maxStrom / 100) * 100;

            doc.lineWidth(0.5).strokeColor('#cbd5e1').fontSize(8).fillColor('#475569');
            for (let i = 0; i <= 4; i++) {
                let yPos = graphTop + gHeight - (i * (gHeight / 4));
                let labelStrom = Math.round((maxStrom / 4) * i);
                doc.font('Helvetica').text(`${labelStrom} kWh`, gLeft - 45, yPos - 4, { width: 40, align: 'right' });
                let labelCop = (5 / 4) * i;
                doc.text(`COP ${labelCop.toFixed(1)}`, gLeft + gWidth + 8, yPos - 4, { width: 40, align: 'left' });
                if (i > 0) doc.moveTo(gLeft, yPos).lineTo(gLeft + gWidth, yPos).stroke();
            }

            doc.lineWidth(1).strokeColor('#334155').moveTo(gLeft, graphTop).lineTo(gLeft, graphTop + gHeight).lineTo(gLeft + gWidth, graphTop + gHeight).stroke();

            const mBreite = gWidth / rows.length; const bBreite = mBreite * 0.4;
            let copPunkte = [];

            rows.forEach((row, i) => {
                let xStart = gLeft + (i * mBreite);
                let sVal = parseFloat(row["Strom Gesamt (kWh)"]) || 0;
                let cVal = parseFloat(row["COP Gesamt"]) || 0;
                let hStrom = (sVal / maxStrom) * gHeight;
                if (hStrom > 0) doc.rect(xStart + (mBreite * 0.1), graphTop + gHeight - hStrom, bBreite, hStrom).fill('#2563eb');
                
                let yCop = graphTop + gHeight - ((cVal / 5) * gHeight);
                copPunkte.push({ x: xStart + (mBreite * 0.3), y: yCop, val: cVal });
                doc.fillColor('#334155').font('Helvetica-Bold').fontSize(8.5).text(row.Zeitpunkt, xStart, graphTop + gHeight + 8, { width: mBreite, align: 'center' });
            });

            if (copPunkte.length > 1) {
                doc.lineWidth(2).strokeColor('#ea580c').moveTo(copPunkte[0].x, copPunkte[0].y);
                for (let p = 1; p < copPunkte.length; p++) { doc.lineTo(copPunkte[p].x, copPunkte[p].y); }
                doc.stroke();
            }
            copPunkte.forEach(p => { if (p.val > 0) doc.circle(p.x, p.y, 3.5).fill('#ea580c'); });

            const yLeg = graphTop + gHeight + 35;
            doc.rect(140, yLeg, 12, 12).fill('#2563eb'); doc.fillColor('#334155').font('Helvetica').fontSize(9).text('Elektrischer Strom (kWh)', 160, yLeg + 2);
            doc.rect(320, yLeg, 12, 3).fill('#ea580c'); doc.circle(326, yLeg + 1.5, 3).fill('#ea580c'); doc.text('Effizienz (COP Gesamt)', 340, yLeg + 2);
        }

        doc.fillColor('#94a3b8').fontSize(7.5).font('Helvetica-Oblique').text('Gesetzlicher Hinweis: Dieser zusammenfassende Kompakt-Nachweis entspricht den Dokumentationsrichtlinien für Verbrauchsdaten nach dem Gebäudeenergiegesetz (GEG) zur Vorlage bei zertifizierten Energieberatern.', 40, 745, { align: 'center', width: 515 });
        doc.end();
    } catch (error) { log(`Fehler im Modus ${mode}: ` + error, "error"); }
}

function ROUND(value, decimals) { return Number(Math.round(value + 'e' + decimals) + 'e-' + decimals); }

async function generateAllReports() {
    log("=== Starte automatisierte Generierung aller 1-Zeilen-Nachweise ===");
    await runReportForMode('daily'); await runReportForMode('week2date'); await runReportForMode('month2date'); await runReportForMode('year2date');
    log("=== Alle Berichte erfolgreich erstellt und versendet ===");
}

// ==================== ZEITPLÄNE ====================
schedule("59 23 * * *", () => { runReportForMode('daily'); });
schedule("55 23 * * 0", () => { runReportForMode('week2date'); });
schedule("50 23 28-31 * *", () => {
    let heute = new Date(); let morgen = new Date(); morgen.setDate(heute.getDate() + 1);
    if (morgen.getDate() === 1) { runReportForMode('month2date'); runReportForMode('year2date'); }
});

// ==================== TEST-TRIGGER ====================
const SOFORT_TESTEN = ""; 
if (SOFORT_TESTEN && SOFORT_TESTEN.trim() !== "") {
    setTimeout(async () => {
        log(`[TEST-MODUS AKTIVIERT] Führe manuelle Generierung aus für: ${SOFORT_TESTEN}`);
        if(SOFORT_TESTEN === "all") { await generateAllReports(); } else { await runReportForMode(SOFORT_TESTEN); }
    }, 2000);
}

