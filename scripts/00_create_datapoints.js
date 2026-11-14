/**
 * ============================================================================
 * Titel:        ioBroker Datenpunkt-Setup (Novelan / Luxtronic Zähler)
 * Projekt:      ioBroker Novelan/Luxtronic – PostgreSQL Energie- und Wärmemengenzähler
 * Version:      1.0
 * Beschreibung: Legt einmalig alle benötigten '0_userdata.0'-Datenpunkte
 *               für Strompreis, elektrische Leistungsaufnahme (kWh) und
 *               erzeugte Wärmemengen (kWh) im ioBroker an.
 *
 * Ausführung:   Einmalig starten – alle fehlenden Datenpunkte werden erstellt,
 *               bereits vorhandene bleiben unangetastet. Danach stoppt sich
 *               das Skript automatisch selbst.
 * ============================================================================
 */

const datapointsToCreate = [
    // Strompreis
    {
        id: '0_userdata.0.energy.contract.electricity.kWhPrice',
        name: 'Aktueller Strompreis pro kWh',
        type: 'number',
        role: 'value.price',
        unit: '€/kWh',
        def: 0.30
    },
    // Elektrische Leistungsaufnahme (kWh)
    {
        id: '0_userdata.0.Wärmepumpe.Leistungsaufnahme.Gesamt',
        name: 'Wärmepumpe Stromverbrauch Gesamt',
        type: 'number',
        role: 'value.energy',
        unit: 'kWh',
        def: 0
    },
    {
        id: '0_userdata.0.Wärmepumpe.Leistungsaufnahme.Heizung',
        name: 'Wärmepumpe Stromverbrauch Heizung',
        type: 'number',
        role: 'value.energy',
        unit: 'kWh',
        def: 0
    },
    {
        id: '0_userdata.0.Wärmepumpe.Leistungsaufnahme.Warmwasser',
        name: 'Wärmepumpe Stromverbrauch Warmwasser',
        type: 'number',
        role: 'value.energy',
        unit: 'kWh',
        def: 0
    },
    // Erzeugte Wärmemenge (kWh)
    {
        id: '0_userdata.0.Wärmepumpe.Wärmemenge.Gesamt',
        name: 'Wärmepumpe Wärmemenge Gesamt',
        type: 'number',
        role: 'value.energy',
        unit: 'kWh',
        def: 0
    },
    {
        id: '0_userdata.0.Wärmepumpe.Wärmemenge.Heizung',
        name: 'Wärmepumpe Wärmemenge Heizung',
        type: 'number',
        role: 'value.energy',
        unit: 'kWh',
        def: 0
    },
    {
        id: '0_userdata.0.Wärmepumpe.Wärmemenge.Warmwasser',
        name: 'Wärmepumpe Wärmemenge Warmwasser',
        type: 'number',
        role: 'value.energy',
        unit: 'kWh',
        def: 0
    }
];

async function createAllDatapoints() {
    log('=== Starte Anlage der ioBroker Datenpunkte ===');

    for (const dp of datapointsToCreate) {
        if (!existsState(dp.id)) {
            await createStateAsync(dp.id, dp.def, {
                name: dp.name,
                type: dp.type,
                role: dp.role,
                unit: dp.unit,
                read: true,
                write: true
            });
            log(`✅ Datenpunkt neu angelegt: ${dp.id} (${dp.unit})`);
        } else {
            log(`ℹ️ Datenpunkt existiert bereits (übersprungen): ${dp.id}`);
        }
    }

    log('=== Datenpunkt-Setup abgeschlossen. Skript wird beendet. ===');
    stopScript();
}

createAllDatapoints();
