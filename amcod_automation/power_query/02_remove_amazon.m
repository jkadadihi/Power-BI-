// Power Query M: Raw Data No Amazon
//
// Replaces AMCOD Stage 5 ("Remove Amazon"): AMCOD reporting excludes Amazon,
// so Marcia filters the raw data, removes Amazon rows, and builds a
// separate "Raw Data No Amazon" dataset by hand each month.
//
// This is a deterministic rule (Priority 2 in the automation plan) - a
// straight filter, no judgment calls - which makes it one of the safest
// and easiest wins to automate.
//
// Reference this query from RawDataEU (01_debits_to_raw_data_eu.m output,
// or the existing Raw Data EU tab) rather than re-reading the source files.
// If "Amazon" ever appears under a different legal-entity name in the
// customer field (e.g. "Amazon EU Sarl", "Amazon Fulfillment"), add it to
// the AmazonNames list below rather than writing a new filter.

let
    Source = RawDataEU,

    AmazonNames = {
        "Amazon", "Amazon EU", "Amazon EU Sarl", "Amazon Fulfillment",
        "Amazon Europe Core", "Amazon Web Services", "Amazon Media EU"
    },

    RemovedAmazon = Table.SelectRows(Source,
        each not List.AnyTrue(List.Transform(AmazonNames,
            (name) => Text.Contains(Text.Upper([customer]), Text.Upper(name))))
        )
in
    RemovedAmazon
