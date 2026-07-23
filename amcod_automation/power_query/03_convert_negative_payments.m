// Power Query M: Payment Sign Cleanup
//
// Replaces AMCOD Stage 4 ("Payment Cleanup"): SAP exports payments as
// negatives (e.g. -500,000); reporting should show them as positives, so
// Marcia manually flips the sign on every payment row each month.
//
// Priority 3 in the automation plan - a simple, mechanical transform that
// removes one more place a manual edit can be forgotten or fat-fingered.
//
// Every negative amount_eur is flipped to positive, matching the
// "-500,000 becomes 500,000" example from the process walkthrough. If only
// rows marked as payments (e.g. debit_credit = "Payment") should be
// flipped, and other negative debit/credit types must stay negative, add
// an `if [debit_credit] = "Payment" then ... else [amount_eur]` guard
// inside the transform below.

let
    Source = RawDataNoAmazon,

    FixedSign = Table.TransformColumns(Source, {
        {"amount_eur", each if _ < 0 then 0 - _ else _, Currency.Type}
    })
in
    FixedSign
