CREATE TABLE `main-dev-431619.demo.data`
(
  id STRING,
  name STRING,
  ingestion_dt DATE,
  ingestion_ts TIMESTAMP
) PARTITION BY ingestion_dt