AED_RATE = 3.68
EUR_RATE = 0.92
QAR_RATE = 3.64
SAR_RATE = 3.75

CURRENCY_CONVERSION_RATES = {
    "USD": 1.0,
    "AED": AED_RATE,
    "EUR": EUR_RATE,
    "SAR": SAR_RATE,
    "QAR": QAR_RATE,
}

CURRENCY_NUMBER_FORMATS = {
    "USD": '"$"#,##0.00',
    "AED": '"AED" #,##0.00',
    "EUR": '"€"#,##0.00',
    "SAR": '"SAR" #,##0.00',
    "QAR": '"QAR" #,##0.00',
}


UAE_TERMS_TITLE = "Terms and Conditions"

# UAE (AED) quotation terms, as provided by Mindware.
UAE_TERMS_AND_CONDITIONS = [
    '1. Payment Terms. Payment shall be made in accordance with the payment terms specified in this quotation.',
    "2. Purchase Orders. All Purchase Orders must be issued to Mindware Technology Trading LLC and must reference this quotation. Purchase Orders shall be duly signed and stamped by the Partner, unless electronically generated through the Partner's approved procurement system.",
    '3. End User Identification. The Partner shall clearly specify the name of the end user/customer in the Purchase Order.',
    '4. Non-Cancellable and Non-Returnable Orders. All accepted orders are firm, non-cancellable, and non-returnable ("NCNR"), unless otherwise expressly agreed in writing by Mindware Technology Trading LLC.',
    '5. Delivery. The delivery of the Products under this quotation is limited to 4-12 weeks, unless otherwise expressly stated herein.',
    '6. Pricing and Taxes. All prices are quoted in AED and are exclusive of VAT, withholding taxes, customs duties, and any other applicable taxes, levies, or governmental charges, which shall be borne by the Partner.',
    '7. Price Validity. Prices contained in this quotation are subject to change at any time prior to receipt and acceptance by Mindware Technology Trading LLC of a valid Purchase Order.',
    '8. Delivery. The estimated delivery lead time is 4-12 weeks and is subject to product availability, manufacturer/vendor confirmation, logistics constraints, governmental approvals, export/import regulations, and any other applicable requirements.',
    '9. Subscription and License Renewals. Unless otherwise specified by the manufacturer/vendor, all subscription and license renewals shall commence from the expiration date of the existing subscription or license term.',
    '10. Automatic Renewal of Subscription and Licenses. Unless otherwise specified by the applicable manufacturer or vendor, any subscription, maintenance, support, or license renewal quoted by Mindware Technology Trading LLC shall automatically renew for the applicable renewal term unless the Partner provides written notice of its intention not to renew at least ninety (90) days prior to the expiration of the then-current subscription, maintenance, support, or license term. In the absence of such notice, the renewal shall be deemed accepted by the Partner, and the Partner shall remain liable for all applicable renewal fees and charges.',
    '11. Quantity Limitation. The pricing offered under this quotation is valid solely for the quantities expressly specified herein. Any change in quantities may result in revised pricing.',
    '12. Warranty. Products, software, and services supplied under this quotation are subject solely to the applicable manufacturer’s or vendor’s warranty and support policies. Mindware Technology Trading LLC provides no additional warranties, whether express, implied, statutory, or otherwise.',
    '13. Exclusion of Commercial Remedies. No cancellation charges, penalties, liquidated damages, service credits, retention amounts, back charges, set-offs, deductions, or similar remedies shall apply to Mindware Technology Trading LLC in connection with any accepted Purchase Order, unless expressly agreed in writing by Mindware Technology Trading LLC.',
    '14. Force Majeure and Supply Chain Disruptions. The Partner acknowledges that regional geopolitical developments, including those affecting the Middle East, may impact international transportation routes, customs clearance processes, supply chains, and product availability. Such events, together with any other events beyond the reasonable control of Mindware Technology Trading LLC, shall constitute Force Majeure events. Mindware Technology Trading LLC shall not be liable for any delay, non-performance, increased costs, or failure to deliver resulting from such events.',
    '15. Governing Terms. The terms and conditions contained in this quotation shall govern and form an integral part of any Purchase Order issued by the Partner pursuant to or in connection with this quotation and any related transaction between Mindware Technology Trading LLC and the Partner, notwithstanding the expiration of the quotation validity period, unless otherwise expressly agreed in writing by Mindware Technology Trading LLC.',
    '16. Conflicting Terms. Any additional, conflicting, inconsistent, or supplementary terms contained in any Purchase Order, acknowledgment, acceptance, procurement portal, standard terms and conditions, or other document issued by the Partner shall be null and void and shall not apply unless expressly accepted in writing by an authorized representative of Mindware Technology Trading LLC.',
    '17. Order Acceptance. No Purchase Order shall be binding upon Mindware Technology Trading LLC unless and until accepted in writing by Mindware Technology Trading LLC.',
    '18. Order of Precedence. In the event of any conflict or inconsistency between the terms of this quotation and any document issued by the Partner, the terms of this quotation shall prevail.',
]


def get_currency_rate(currency_code: str) -> float:
    return CURRENCY_CONVERSION_RATES.get((currency_code or "AED").upper(), AED_RATE)


def get_currency_format(currency_code: str) -> str:
    return CURRENCY_NUMBER_FORMATS.get((currency_code or "AED").upper(), '"AED" #,##0.00')


def get_footer_notes(currency_code: str):
    code = (currency_code or "AED").upper()
    if code == "AED":
        return [UAE_TERMS_TITLE] + list(UAE_TERMS_AND_CONDITIONS)
    if code == "AED_LEGACY":
        return [
            "Ø  Payment terms will be as per our finance approval.",
            "Ø  These prices are till DDP Dubai.",
            "Ø  Hardware will take 4-12 weeks delivery time from the date of Booking.",
            "Ø  These prices do not include Mindware installation of any kind.",
            "Ø  Change in Qty or partial shipment is not acceptable.",
            f"Ø  PO Should be addressed to Mindware Technology Trading LLC and should be in {code}.",
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
