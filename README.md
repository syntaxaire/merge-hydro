# Split Hydro One bills for further processing
Ontario's Hydro One electric utility provides detailed billing in Excel format. These bills
are difficult to work with since they provide only one month's data at a time and can be unpredictably
formatted.

This script takes in an arbitrary number of Hydro One bills in .xls format and outputs an `output.xls` with
the results collated into one worksheet per account/meter.

This should also work if you only have one meter, making it easier to graph your power usage.

Some configuration options are found in `process.cfg`. Copy `process.cfg.example` to `process.cfg` to make it work. 