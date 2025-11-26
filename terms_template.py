def get_terms_section(header_info, total_price_sum):
    quote_validity = header_info.get("Bid Expiration Date", "N/A")
    company_name = header_info.get("Reseller Name", "Company")
    end_user = header_info.get("Customer Name", "End User")
    formatted_price = f"AED {total_price_sum:,.2f}"
    return [
        ("B29", "Terms and Conditions:", {"bold": True, "size": 11, "color": "1F497D"}),
        # Main terms paragraph
        ("C30", f"""THIS DOCUMENT WILL BE GOVERNED BY THE TERMS AND CONDITIONS MENTIONED BELOW, THE AGREEMENT ENTERED INTO WITH MINDWARE AND ANY ASSOCIATED IBM TERMS AND CONDITIONS, AS APPLICABLE.
• 30 Days from POE Date.
• Quote Validity: {quote_validity}
• Customer Price should not exceed {formatted_price}
• Pricing valid for this transaction only."""
        ),
        # Additional terms
        ("C31", """• Business Partner has to clearly understand IBM’s Terms and Conditions governing Software and Software as a Service offerings and comply with their contractual requirements at the time of placing the order (Example: Business Partner Agreement for SVP, IBM Cloud Offerings).
• Business Partner will notify Mindware & IBM in case there is a Non-Disclosure Agreement between Business Partner and the End User.
• Prices quoted are End-User Price in US Dollars. As per IBM compliance rules, Business Partner understands that he/she cannot resell to the End User at a price higher than the one quoted above. Business Partner’s quotation and invoice to the End User customer must mention the part number, description, quantity, prices as per above.
• IBM Program incentives are only eligible if the opportunity meets all the programs requirement as captured in Programs Operations Guide. For more info, follow the link below: http://www.ibm.biz/PartnerPlusIncentiveProgramBPs. If IBM rejects the incentive because of any insufficient requirement/criteria, Mindware will revoke the incentive that is passed and revert to normal margin.
• Business Partner is responsible to ensure that the above Bill of Material is in line with the End User requirements. No order cancellation will be accepted once your PO is received.
• Documents & Media shipping, insurance, clearing and customs duties not included
• Applicable duties, charges & taxes not included
• Attached quote does not include installation services and training
• Conversion Rate $1 US = 3.6725 AED"""
        ),

        # Compliance statement
        ("C32", f"""Please include the following phrase as per IBM Compliance within your PO:
This is to confirm that {company_name} has accepted a firm and final order from {end_user} 
and the Final end-customer price for IBM SW licenses is not higher than {formatted_price}"""
        ),
        ("C33", f"""This documents confirms {company_name} commitment to place a firm order on IBM through Mindware FZ-LLC consistent with the End Customer Purchase Order (PO) or any other legally binding document {company_name} received from the end customer.
"""
        ),
    
        ("C34", "IBM’s Review of Your Compliance with this Agreement", {"bold": True}),
        ("C35", """IBM may periodically review your compliance with this Agreement. You agree to provide IBM with relevant records on request. IBM may reproduce and retain copies of these records. IBM, or an independent auditor, may conduct a review of your compliance with this Agreement on your premises during your normal business hours."""
        ),
        ("C36", """If, during the review of your compliance with this Agreement, it is determined you have failed to comply with any material term of this Agreement, in addition to rights under law and the terms of this Agreement, for transactions that are the subject of the breach, you agree to refund the amount equal to the discount (or fee, if applicable) IBM gave you for the applicable Products or Services or IBM may offset any amounts due to you from IBM."""
        ),
        ("C37", "Notice: Re-export of these products may be subject to local & United States Department of Commerce export regulations."),
        ("C38", "It is your responsibility to ensure compliance with all such regulations."),
        ("C39", f"""Compliance Review
Transaction Agreement Reseller (“Reseller”) shall keep and maintain all records necessary to establish its compliance with the Agreement for at least three years after the Agreement end date. IBM and/or VAD or their auditors may periodically review Reseller’s compliance with the Agreement, and may do so either remotely, on Reseller’s premises during normal business hours, or a combination thereof. In connection with any such review, Reseller’s agrees to provide IBM and/or VAD, or their auditor, with relevant records and system tools output on request. IBM and/or VAD may reproduce and retain copies of such records and output.
If, during any such review, it is determined that Reseller has failed to comply with any material term of this Agreement, in addition to IBM’s and or VAD’s rights under law and the terms of this Agreement, for transactions that are the subject of the breach, Reseller agrees to refund the amount equal to the discount or fees, if any, that IBM gave Reseller through VAD for the applicable Products or Services, or IBM and or VAD may offset any amounts due to Reseller from VAD.
IBM’s audit rights with respect to special bids are set forth further in Section 6.
Compliance with Laws
Each party will comply with all laws and regulations applicable to its business and content, including, without limitation, those prohibiting corruption and bribery, such as the U.S. Foreign Corrupt Practices Act and those governing transactions with government and public entities, antitrust and competition, taxes and export insider trading, securities, and financial reporting, consumer transactions, and regarding data privacy. Each party will procure all licenses and pay all fees and other charges required for such compliance.
Prohibition of Inappropriate Conduct
Reseller will not directly or indirectly make or give, offer or promise to make or give, or authorize the making or giving of any payment, gift, or other thing of value or advantage (including, for example, accommodations, air fare, entertainment or meals) to any person or entity for (a) the purpose of (i) wrongfully influencing any act or decision, (ii) inducing any act or omission to act in violation of a lawful duty; (iii) inducing a misuse of influence or (iv) securing any improper advantage, or (b) any purpose that is otherwise unlawful under any applicable anti-corruption or anti-bribery law, including the U.S. Foreign Corrupt Practices Act. VAD may terminate this Agreement immediately if Reseller breaches this Section or if VAD reasonably believes such a breach has occurred or is likely to occur.
Code of Conduct
Reseller agrees to comply with the IBM Code of Conduct, a current version of which is available on the following IBM Internet website: 
https://www.ibm.com/investor/att/pdf/IBM_Business_Conduct_Guidelines.pdf
Reseller agrees to comply with the Midis Group Code of Conduct, a current version of which is available on the Midis Group Website: 
https://www.midisgroup.com/wp-content/uploads/2024/08/Code-of-Conduct-2023-English.pdf 
Special Bids
Reseller may request a Special Bid (a special discount or price) on a specific End User transaction. VAD may, at its sole discretion, approve or reject a Special Bid based on the information provided by Reseller in its Special Bid request. If VAD approves a Special Bid, then the price provided by VAD shall only be valid for the applicable Special Bid, and its validity shall be subject to all the terms and conditions set out in this Agreement, including IBM's Special Bid authorization notice (“Special Bid Addendum”). Further, IBM provides Special Bids through VAD to Reseller on the basis that the information Reseller provided in its Special Bid request is truthful and accurate. If the information provided in the Special Bid request changes, Reseller must immediately notify VAD. In such event, VAD reserves the right to modify the terms of, or cancel any Special Bid authorization it may have provided. If Reseller fails to provide truthful and accurate information on Special Bid requests, then VAD shall be entitled to recover from Reseller (and Reseller is obligated to repay) the amount of any discounts IBM provided through VAD in the Special Bid and take any other actions authorized under this Agreement or applicable law. Special Bid authorizations and the terms applicable to Special Bids are IBM’s confidential information, which is subject to the applicable confidentiality agreement.
Reseller accepts the terms of the Special Bid by:
a. submitting an order under the Special Bid authorization;
b. accepting the Products or Services for which Reseller is receiving a Special Bid;
c. providing the Products or Services to its Customer; or
d. paying for the Products or Services.
The Special Bid discount or price for eligible Products or Services is subject to the following:
a. no other discounts, incentive offerings, rebates, or promotions apply, unless VAD specifies otherwise in writing;
b. availability of the Products or Services;
c. Reseller’s acceptance of the additional terms contained in the Special Bid Addendum (which occurs upon Reseller’s acceptance of the Special Bid, as set forth above)
d. Reseller’s advising the local VAD financing entity/organization of any Special Bid pricing for any Products or Services for which Reseller arranges financing; and
e. Resale of the Products or Services by Reseller to the End User associated with the Special Bid by the date indicated in the Special Bid request.
If reseller is a Distributor, Reseller may only market the Products and Services to the Resellers that Reseller has identified in the Special Bid request as bidding to the End User.
Reseller is responsible to require Reseller’s Resellers who do not have a contract with IBM to market such Products and Services to comply with the Special Bid terms contained in this Agreement and in the applicable Special Bid Addendum that IBM provides for the Special Bid through VAD.
If Reseller is requesting a specific End User price or discount in the Special Bid, Reseller shall ensure, and shall require any Resellers to also ensure, that the intended End User receives the financial benefit of such price or discount.
IBM’s Audit of Special Bid Transactions
IBM may audit directly or through VAD any Special Bid transactions in accordance with the terms of this Section 
a. Upon VAD’s request, Reseller will promptly provide VAD or its auditors with all relevant Documentation to enable VAD and/or IBM to verify that all information provided in support of a Special Bid request was truthful and accurate and that IBM Products and Services have been or will be supplied to the End User in accordance with the terms of the Special Bid, including, but not limited to, i) documentation that identifies the dates of sale and delivery and End User prices for IBM Products and Services, such as invoices, delivery orders, contracts and purchase orders by and between Reseller and any Reseller and by and between Reseller or any Reseller and an End User and ii) documentation that demonstrates that Reseller or Reseller’s Reseller, as applicable, own and use the Special Bid Products for at least the Service Period to provide the service offering described in the terms of the Special Bid to End Users (collectively, items i) and ii) being the “Documentation”).
b. In any case where reseller is unable to provide the Documentation because of confidentiality obligations owed to an End User, whether arising by written contract or applicable law, Reseller will promptly provide VAD with written evidence of, and any Documentation not subject to, those obligations. In addition, Reseller will promptly and in writing seek the End User’s consent to waive confidentiality restrictions to permit VAD and IBM to conduct their audit as intended. Should the End User refuse to grant that consent, Reseller will i) provide VAD with a copy of the waiver request and written proof of that refusal and ii) identify appropriate contacts at the End User with whom VAD may elect to discuss the refusal.
c. Reseller hereby waives any objection to i) VAD and/or IBM sharing Special Bid information directly with the End User, notwithstanding the terms of any agreement that would prohibit VAD from doing so, and otherwise communicating (both orally and in writing) with the End User, as VAD deems necessary and appropriate to complete its desired compliance review and ii) the End User sharing Special Bid information directly with IBM/VAD. In this subsection (c), “Special Bid information” includes, but is not limited to, the types and quantity of Products and anticipated End User prices and delivery dates set forth in a Special Bid. VAD may invalidate a Special Bid if in respect of such Special Bid, Reseller fails to comply with this Section 4.9.1 or the applicable Special Bid terms. In that event, IBM/VAD shall be entitled to recover from Reseller (and Reseller is obligated to repay) the amount of any discounts IBM provided in the Special Bid. IBM may also take any other actions authorized under the Agreement or applicable law.""", {"merge_rows": 2}),
        ("C41", f"""These Terms and Conditions are binding and form an integral part of any agreement between Mindware FZ-LLC and the Partner.
Mindware may include additional terms in its quotations before its acceptance by the Partner, which the Partner agrees to be bound by as if fully incorporated into the agreement.
Notes:
It is understood that final coverage dates for New Licenses and Software Subscription and Support Reinstatement part numbers will be based upon IBM's acceptance of the Purchase Order, as defined by the Passport Advantage program, irrespective of the dates which may appear in this special bid.
Pricing is valid for this transaction only.
For Future Offers, any margin discounts, incentives, or promotions will be applied based on then current availability and terms at the actual time of billing for the Future Offer. For clarity, this means that the pricing for any Future Offering in this Quotation will be recalculated at the time of Distributor's actual order to apply existing applicable margin discounts, incentives, and promotions in effect at that time.
This quote is contingent upon all paperwork being on file with IBM.
Thank you for returning this quotation duly signed and stamped for your approval.
Signature                          …………………………..
Name                               ………………………………
Title                              …………………………………
Date                               ………………………………...""", {"merge_rows": 5})
    ]
