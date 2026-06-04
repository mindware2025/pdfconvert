AED_RATE = 3.68
EUR_RATE = 0.92
QAR_RATE = 3.64

CURRENCY_CONVERSION_RATES = {
    "USD": 1.0,
    "AED": AED_RATE,
    "EUR": EUR_RATE,
    "QAR": QAR_RATE,
}

CURRENCY_NUMBER_FORMATS = {
    "USD": '"$"#,##0.00',
    "AED": '"AED" #,##0.00',
    "EUR": '"€"#,##0.00',
    "QAR": '"QAR" #,##0.00',
}


def get_currency_rate(currency_code: str) -> float:
    return CURRENCY_CONVERSION_RATES.get((currency_code or "AED").upper(), AED_RATE)


def get_currency_format(currency_code: str) -> str:
    return CURRENCY_NUMBER_FORMATS.get((currency_code or "AED").upper(), '"AED" #,##0.00')


def get_footer_notes(currency_code: str):
    code = (currency_code or "AED").upper()
    if code == "AED":
        return [
            "Ø  Payment terms will be as per our finance approval.",
            "Ø  These prices are till DDP Dubai.",
            "Ø  Hardware will take 4-12 weeks delivery time from the date of Booking.",
            "Ø  These prices do not include Mindware installation of any kind.",
            "Ø  Change in Qty or partial shipment is not acceptable.",
            "Ø  PO Should be addressed to Mindware Technology Trading LLC and should be in AED.",
            "Ø  For all B2B orders complete end customer details should be mentioned on the PO.",
            "Ø  Orders once placed with Dell cannot be cancelled.",
            "Ø  Kindly also ensure to review the proposal specifications from your end and ensure that they match the requirements exactly as per the End User.",
            "Ø  Partial deliveries shall be acceptable",
            "Ø  For UAE DDP orders, the PO should be addressed to Mindware Technology Trading LLC and for Ex-Jablal Ali orders, it should be addressed to Mindware FZ.",
            "Ø  Please ensure that the PO includes the name of the end-user.",
            "Ø  Please ensure that the PO includes the Incoterms (DDP or Ex-Works Jabal Ali).",
            "Ø  Due to global market fluctuations, all prices are subject to change without prior notice, and lead times may also be affected. All quotations are non-binding and remain subject to final validation and confirmation by Dell.",
            "Ø  As the geopolitical situation in the Middle East continues to evolve, it has introduced significant instability to international shipping routes. These unforeseen and extraordinary circumstances, which remain entirely beyond our control, constitute a Force Majeure event. We are formally notifying you of the resulting impact on our current and future shipments.",
        ]

    return [
        "Incoterms:",
        "",
        "Payment Terms:",
        "",
        "Quote validity:",
        "",
        "Estimated Delivery Time from the date of booking:",
        "",
        "These prices do not include installation of any kind",
        "All prices are exclusive of VAT and any other applicable taxes, which shall be charged in accordance with applicable laws and regulations.",
        "Change in Qty or partial shipment is not acceptable",
        "For all B2B orders complete end customer details should be mentioned on the PO",
        f"PO Should be addressed to Mindware FZ LLC and should be in {code}",
        "Orders once placed with Dell cannot be cancelled",
    ]
