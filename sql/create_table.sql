-- ============================================================================
-- PostgreSQL Tabellen-Setup für Novelan / Luxtronic Wärmepumpen-Datenlogger
-- Datenbank: smarthome | Tabelle: heat_pump
-- ============================================================================

CREATE TABLE IF NOT EXISTS public.heat_pump (
    id SERIAL PRIMARY KEY,
    "timestamp" TIMESTAMPTZ NOT NULL DEFAULT now(),
    "priceIn" DOUBLE PRECISION,
    "priceOut" DOUBLE PRECISION,
    "GesamtLeistung-kWh" DOUBLE PRECISION,
    "Heizung-kWh" DOUBLE PRECISION,
    "Warmwasser-kWh" DOUBLE PRECISION,
    "GesamteWärmenMengekWh" DOUBLE PRECISION,
    "Heizung-WM-kWh" DOUBLE PRECISION,
    "Warmwasser-WM-kWh" DOUBLE PRECISION
);

-- Index auf den Zeitstempel für schnelle Tages-, Monats- und Jahresabfragen
CREATE INDEX IF NOT EXISTS idx_heat_pump_timestamp ON public.heat_pump ("timestamp" DESC);
CREATE EXTENSION IF NOT EXISTS timescaledb CASCADE;
SELECT create_hypertable('heat_pump', 'timestamp', if_not_exists => TRUE);
