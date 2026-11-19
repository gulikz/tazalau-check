/* =====================================================================================================
   hes_pkb_monitoring_agg.sql  (Vertica)

   PKB (Первое кредитное бюро) monitoring events -> per client (IIN) x rule aggregates, LONG format.

   Source : dds.hes_pkb_monitoring_event   (~110M rows, ~1.9M IINs, events since 2026-09-07)
   Output : one row per (iin, rule_key, rule_name):
              iin            12-char IIN, leading zeros kept
              rule_key       e.g. 'CustomRule.CMC'
              rule_name      e.g. 'Мониторинг статуса «Лудоман»' (the notebook maps names -> codes)
              cnt_7d         events in [snapshot_date - 7 days,  snapshot_date)
              cnt_30d        events in [snapshot_date - 30 days, snapshot_date)
              cnt_60d        events in [snapshot_date - 60 days, snapshot_date)
              cnt_all        all events before snapshot_date (scan bounded to 400 days)
              first_ts       first event time
              last_ts        last event time
              snapshot_date  the parameter (the aggregate is "as of" this date, no look-ahead)
   Use    : save the result as parquet  ->  CONFIG['PKB_AGG_FILE'] (default DATA_DIR/hes_pkb_agg.parquet).
            The notebook block fb_pkb_monitoring uses it for OVERLAY rules only (history is too short to
            train on); once 6+ months have accumulated, build monthly snapshots and move it into the model.

   Parameter
     :snapshot_date   DATE literal, e.g. '2026-11-18'. Only events strictly BEFORE this date are counted.
                      Python: sql.replace(':snapshot_date', "'2026-11-18'")
                      vsql  : \set snapshot_date '''2026-11-18'''

   >>> PLACEHOLDERS: the IIN and event-timestamp column names were not visible in the screenshots. <<<
     iin_col__PLACEHOLDER       -> the client IIN column        (check: SELECT * FROM dds.hes_pkb_monitoring_event LIMIT 5;)
     event_ts_col__PLACEHOLDER  -> the event date-time column   (the one spanning 2026-09-07 .. today)
   Replace every occurrence (search for "__PLACEHOLDER") before running.
   ===================================================================================================== */
WITH params AS (
    SELECT CAST(:snapshot_date AS TIMESTAMP) AS snap_ts
),
ev AS (
    SELECT
        -- PLACEHOLDER column: client IIN -> digits only, left-padded to 12 characters
        LPAD(REGEXP_REPLACE(CAST(e.iin_col__PLACEHOLDER AS VARCHAR(32)), '[^0-9]', ''), 12, '0') AS iin,
        -- PLACEHOLDER column: event timestamp
        CAST(e.event_ts_col__PLACEHOLDER AS TIMESTAMP)                                         AS event_ts,
        e.rule_key,
        e.rule_name
    FROM dds.hes_pkb_monitoring_event e
    CROSS JOIN params p
    WHERE e.event_ts_col__PLACEHOLDER <  p.snap_ts                         -- PLACEHOLDER column: strictly before the snapshot
      AND e.event_ts_col__PLACEHOLDER >= p.snap_ts - INTERVAL '400 days'   -- PLACEHOLDER column: bounded scan
      AND e.rule_name NOT LIKE 'Отписка%'                                  -- technical rule (unsubscribe), not a client event
      -- Optional: keep only the auto-loan portfolio to shrink the output
      -- AND e.iin_col__PLACEHOLDER IN (SELECT iin FROM <your_portfolio_table>)
)
SELECT
    ev.iin,
    ev.rule_key,
    ev.rule_name,
    SUM(CASE WHEN ev.event_ts >= p.snap_ts - INTERVAL '7 days'  THEN 1 ELSE 0 END) AS cnt_7d,
    SUM(CASE WHEN ev.event_ts >= p.snap_ts - INTERVAL '30 days' THEN 1 ELSE 0 END) AS cnt_30d,
    SUM(CASE WHEN ev.event_ts >= p.snap_ts - INTERVAL '60 days' THEN 1 ELSE 0 END) AS cnt_60d,
    COUNT(*)                                                                      AS cnt_all,
    MIN(ev.event_ts)                                                              AS first_ts,
    MAX(ev.event_ts)                                                              AS last_ts,
    CAST(p.snap_ts AS DATE)                                                       AS snapshot_date
FROM ev
CROSS JOIN params p
GROUP BY ev.iin, ev.rule_key, ev.rule_name, p.snap_ts
;
