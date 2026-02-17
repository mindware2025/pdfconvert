
def get_terms_section(header_info, total_price_sum):
        # DEBUG: Print all header_info keys and values to diagnose extraction issues
    print("[DEBUG] header_info keys:", list(header_info.keys()))
    print("[DEBUG] header_info values:", header_info)
    quote_validity = header_info.get("Bid Expiration Date", "N/A")
    company_name = header_info.get("Reseller Name", "Company")
    end_user = header_info.get("Customer Name", "End User")
    # Use MEP from PDF if available, otherwise fallback to calculated total
    mep_value = header_info.get("Maximum End User Price (MEP)", "")
    if not mep_value:
         mep_value = header_info.get("Total Value Seller Revenue Opportunity", "")
    # Conversion rate and currency by country: KSA -> 3.75 SAR; UAE/Qatar -> 3.6725 AED
    c = (header_info.get('country') or '').strip().upper()
    rate = 3.75 if c == 'KSA' else 3.6725
    currency = 'SAR' if c == 'KSA' else 'AED'
    # Create text with MEP placeholder - will be processed to include formula
    if mep_value:
                # Store MEP value in header_info for formula use in Excel generation
                try:
                        mep_numeric = float(mep_value.replace(",", ""))
                        header_info["_MEP_NUMERIC"] = mep_numeric
                        mep_local = mep_numeric * rate
                        if header_info.get('country', '').lower() == 'qatar':
                                formatted_price = f"USD {mep_value}"
                        else:
                                formatted_price = f"USD {mep_value} ({currency} {mep_local:,.2f})"
                except:
                        formatted_price = f"USD {mep_value}"
    else:
        formatted_price = f" "
    
    return [
        ("B29", "Terms and Conditions:", {"bold": True, "size": 11, "color": "1F497D"}),
        # Main terms paragraph
        ("C30", f"""THIS DOCUMENT WILL BE GOVERNED BY THE TERMS AND CONDITIONS MENTIONED BELOW, THE AGREEMENT ENTERED INTO WITH MINDWARE AND ANY ASSOCIATED IBM TERMS AND CONDITIONS, AS APPLICABLE.
• Payment Terms: As aligned with Mindware
• Quote Validity: {quote_validity}
• Customer Price should not exceed {formatted_price}
• Pricing valid for this transaction only."""
        ),
        # Additional terms (Qatar: no conversion line; UAE/KSA: show conversion rate)
        ("C31", (
            f"""• Business Partner has to clearly understand IBM's Terms and Conditions governing Software and Software as a Service offerings and comply with their contractual requirements at the time of placing the order (Example: Business Partner Agreement for SVP, IBM Cloud Offerings).
• Business Partner will notify Mindware & IBM in case there is a Non-Disclosure Agreement between Business Partner and the End User.
• Prices quoted are End-User Price in US Dollars. As per IBM compliance rules, Business Partner understands that he/she cannot resell to the End User at a price higher than the one quoted above. Business Partner's quotation and invoice to the End User customer must mention the part number, description, quantity, prices as per above.
• IBM Program incentives are only eligible if the opportunity meets all the programs requirement as captured in Programs Operations Guide. For more info, follow the link below: http://www.ibm.biz/PartnerPlusIncentiveProgramBPs. If IBM rejects the incentive because of any insufficient requirement/criteria, Mindware will revoke the incentive that is passed and revert to normal margin.
• Business Partner is responsible to ensure that the above Bill of Material is in line with the End User requirements. No order cancellation will be accepted once your PO is received.
• Documents & Media shipping, insurance, clearing and customs duties not included
• Applicable duties, charges & taxes not included
• Attached quote does not include installation services and training"""
            + (f"\n• Conversion Rate $1 US = {rate} {currency}" if c != "QATAR" else "")
        )),
        # Compliance statement
        ("C32", f"""Please include the following phrase as per IBM Compliance within your PO:
This is to confirm that {company_name} has accepted a firm and final order from {end_user} 
and the Final end-customer price for IBM SW licenses is not higher than {formatted_price}"""
        ),
        ("C33", f"""This documents confirms {company_name} commitment to place a firm order on IBM through Mindware FZ-LLC consistent with the End Customer Purchase Order (PO) or any other legally binding document {company_name} received from the end customer.
"""
        ),
    
        ("C34", "IBM's Review of Your Compliance with this Agreement", {"bold": True}),
        ("C35", """IBM may periodically review your compliance with this Agreement. You agree to provide IBM with relevant records on request. IBM may reproduce and retain copies of these records. IBM, or an independent auditor, may conduct a review of your compliance with this Agreement on your premises during your normal business hours."""
        ),
        ("C36", """If, during the review of your compliance with this Agreement, it is determined you have failed to comply with any material term of this Agreement, in addition to rights under law and the terms of this Agreement, for transactions that are the subject of the breach, you agree to refund the amount equal to the discount (or fee, if applicable) IBM gave you for the applicable Products or Services or IBM may offset any amounts due to you from IBM."""
        ),
        ("C37", "Notice: Re-export of these products may be subject to local & United States Department of Commerce export regulations."),
        ("C38", "It is your responsibility to ensure compliance with all such regulations."),
        ("C39", "Compliance Review", {"bold": True}),
        ("C40", """Transaction Agreement Reseller (“Reseller”) shall keep and maintain all records necessary to establish its compliance with the Agreement for at least three years after the Agreement end date."""),
        ("C41", """ IBM and/or VAD or their auditors may periodically review Reseller’s compliance with the Agreement, and may do so either remotely, on Reseller’s premises during normal business hours, or a combination thereof. In connection with any such review, Reseller’s agrees to provide IBM and/or VAD, or their auditor, with relevant records and system tools output on request. IBM and/or VAD may reproduce and retain copies of such records and output."""),
        ("C42", """If, during any such review, it is determined that Reseller has failed to comply with any material term of this Agreement, in addition to IBM’s and or VAD’s rights under law and the terms of this Agreement, for transactions that are the subject of the breach, Reseller agrees to refund the amount equal to the discount or fees, if any, that IBM gave Reseller through VAD for the applicable Products or Services, or IBM and or VAD may offset any amounts due to Reseller from VAD."""),
        ("C43", """IBM's audit rights with respect to special bids are set forth further in Section 6."""),
        ("C44", "Compliance with Laws", {"bold": True}),
        ("C45", """Each party will comply with all laws and regulations applicable to its business and content, including, without limitation, those prohibiting corruption and bribery, such as the U.S. Foreign Corrupt Practices Act and those governing transactions with government and public entities, antitrust and competition, taxes and export insider trading, securities, and financial reporting, consumer transactions, and regarding data privacy. Each party will procure all licenses and pay all fees and other charges required for such compliance."""),
        ("C46", "Prohibition of Inappropriate Conduct", {"bold": True}),
       # ("C47", """Reseller will not directly or indirectly make or give, offer or promise to make or give, or authorize the making or giving of any payment, gift, or other thing of value or advantage for unlawful purposes under applicable anti-corruption or anti-bribery laws."""),
        ("C47", """Reseller will not directly or indirectly make or give, offer or promise to make or give, or authorize the making or giving of any payment, gift, or other thing of value or advantage (including, for example, accommodations, air fare, entertainment or meals) to any person or entity for (a) the purpose of (i) wrongfully influencing any act or decision, (ii) inducing any act or omission to act in violation of a lawful duty; (iii) inducing a misuse of influence or (iv) securing any improper advantage, or (b) any purpose that is otherwise unlawful under any applicable anti-corruption or anti-bribery law, including the U.S. Foreign Corrupt Practices Act. VAD may terminate this Agreement immediately if Reseller breaches this Section or if VAD reasonably believes such a breach has occurred or is likely to occur."""),
        ("C48", "Code of Conduct", {"bold": True}),
        ("C49", """Reseller agrees to comply with the IBM Code of Conduct, a current version of which is available on the following IBM Internet website: 
https://www.ibm.com/investor/att/pdf/IBM_Business_Conduct_Guidelines.pdf
 
Reseller agrees to comply with the Midis Group Code of Conduct, a current version of which is available on the Midis Group Website: 
https://www.midisgroup.com/wp-content/uploads/2024/08/Code-of-Conduct-2023-English.pdf 
"""),
        ("C50", "Special Bids", {"bold": True}),
            ("C51", """Reseller may request a Special Bid (a special discount or price) on a specific End User transaction. VAD may, at its sole discretion, approve or reject a Special Bid based on the information provided by Reseller in its Special Bid request."""),
        
            ("C52", """If VAD approves a Special Bid, then the price provided by VAD shall only be valid for the applicable Special Bid, and its validity shall be subject to all the terms and conditions set out in this Agreement, including IBM's Special Bid authorization notice (“Special Bid Addendum”)."""),
        
            ("C53", """Further, IBM provides Special Bids through VAD to Reseller on the basis that the information Reseller provided in its Special Bid request is truthful and accurate. If the information provided in the Special Bid request changes, Reseller must immediately notify VAD. In such event, VAD reserves the right to modify the terms of, or cancel any Special Bid authorization it may have provided."""),
        
            ("C54", """If Reseller fails to provide truthful and accurate information on Special Bid requests, then VAD shall be entitled to recover from Reseller (and Reseller is obligated to repay) the amount of any discounts IBM provided through VAD in the Special Bid and take any other actions authorized under this Agreement or applicable law."""),
        
            ("C55", """Special Bid authorizations and the terms applicable to Special Bids are IBM’s confidential information, which is subject to the applicable confidentiality agreement."""),
        
            ("C56", """Reseller accepts the terms of the Special Bid by:
        a. submitting an order under the Special Bid authorization;
        b. accepting the Products or Services for which Reseller is receiving a Special Bid;
        c. providing the Products or Services to its Customer; or
        d. paying for the Products or Services."""),
        
            ("C57", """The Special Bid discount or price for eligible Products or Services is subject to the following:
        a. no other discounts, incentive offerings, rebates, or promotions apply, unless VAD specifies otherwise in writing;
        b. availability of the Products or Services;
        c. Reseller’s acceptance of the additional terms contained in the Special Bid Addendum (which occurs upon Reseller’s acceptance of the Special Bid, as set forth above)"""),
        
            ("C58", """d. Reseller’s advising the local VAD financing entity/organization of any Special Bid pricing for any Products or Services for which Reseller arranges financing; and
        e. Resale of the Products or Services by Reseller to the End User associated with the Special Bid by the date indicated in the Special Bid request."""),
        
            ("C59", """If reseller is a Distributor, Reseller may only market the Products and Services to the Resellers that Reseller has identified in the Special Bid request as bidding to the End User."""),
        
            ("C60", """Reseller is responsible to require Reseller’s Resellers who do not have a contract with IBM to market such Products and Services to comply with the Special Bid terms contained in this Agreement and in the applicable Special Bid Addendum that IBM provides for the Special Bid through VAD."""),
        
            ("C61", """If Reseller is requesting a specific End User price or discount in the Special Bid, Reseller shall ensure, and shall require any Resellers to also ensure, that the intended End User receives the financial benefit of such price or discount."""),
        
        
            ("C62", "IBM’s Audit of Special Bid Transactions", {"bold": True}),
        
            ("C63", """IBM may audit directly or through VAD any Special Bid transactions in accordance with the terms of this Section 
        a. Upon VAD’s request, Reseller will promptly provide VAD or its auditors with all relevant Documentation to enable VAD and/or IBM to verify that all information provided in support of a Special Bid request was truthful and accurate and that IBM Products and Services have been or will be supplied to the End User in accordance with the terms of the Special Bid, including, but not limited to, i) documentation that identifies the dates of sale and delivery and End User prices for IBM Products and Services, such as invoices, delivery orders, contracts and purchase orders by and between Reseller and any Reseller and by and between Reseller or any Reseller and an End User and ii) documentation that demonstrates that Reseller or Reseller’s Reseller, as applicable, own and use the Special Bid Products for at least the Service Period to provide the service offering described in the terms of the Special Bid to End Users (collectively, items i) and ii) being the “Documentation”).
        b. In any case where reseller is unable to provide the Documentation because of confidentiality obligations owed to an End User, whether arising by written contract or applicable law, Reseller will promptly provide VAD with written evidence of, and any Documentation not subject to, those obligations. In addition, Reseller will promptly and in writing seek the End User’s consent to waive confidentiality restrictions to permit VAD and IBM to conduct their audit as intended. Should the End User refuse to grant that consent, Reseller will i) provide VAD with a copy of the waiver request and written proof of that refusal and ii) identify appropriate contacts at the End User with whom VAD may elect to discuss the refusal.
        c. Reseller hereby waives any objection to i) VAD and/or IBM sharing Special Bid information directly with the End User, notwithstanding the terms of any agreement that would prohibit VAD from doing so, and otherwise communicating (both orally and in writing) with the End User, as VAD deems necessary and appropriate to complete its desired compliance review and ii) the End User sharing Special Bid information directly with IBM/VAD. In this subsection (c), “Special Bid information” includes, but is not limited to, the types and quantity of Products and anticipated End User prices and delivery dates set forth in a Special Bid. VAD may invalidate a Special Bid if in respect of such Special Bid, Reseller fails to comply with this Section 4.9.1 or the applicable Special Bid terms. In that event, IBM/VAD shall be entitled to recover from Reseller (and Reseller is obligated to repay) the amount of any discounts IBM provided in the Special Bid. IBM may also take any other actions authorized under the Agreement or applicable law."""),
        
            ("C64", """d. For any IBM price offer, the Partner shall ensure that all applicable IBM price offer and Special Bid addendum terms are communicated in writing to any inserted tier involved in the transaction."""),
        
            ("C65", """e. If the Partner retains title to the products, for example for Demonstration, Development, or Managed Service purposes, IBM will consider the IBM Business Partner as the end user and may be required to demonstrate compliance with any specified retention period."""),
        
            ("C66", """f. For any approved solution, the Partner must ensure that products are resold only as part of the approved solution. For Special Bids, the Partner must ensure that the End User is charged no more than the Maximum End User Price (MEP)."""),
        
            ("C67", """g. If the MEP is converted to local currency, the mid-market spot rate must be used, sourced from Bloomberg, Reuters, Wall Street Journal, or the Central Bank within three (3) days prior to proposal submission or contract signature."""),
        
            ("C68", """h. Partners may use hedging strategies to manage currency fluctuation risks. The applicable hedging arrangement must be established with a recognized financial institution (e.g. a bank) prior to the conversion."""),
        
            ("C69", """i. The Partner shall retain records to demonstrate compliance with these requirements upon request by IBM or Mindware."""),
        
            ("C70", """j. For Special Bids, the Partner must ensure that the End User is charged no more than the approved Maximum End User Price (MEP). Ideally, the End User invoice should show IBM products as a separate line item. If the IBM products are part of a bundled price, the Partner must ensure that the pricing structure, additional products, services, or other charges included in the bundle are delivered to the End User and that the End User is aware of them. Standard costs (e.g., installation, warehousing, pre-sales preparation) are included in the approved MEP and cannot be used to justify exceeding the MEP, while other additional costs (e.g., currency hedging fees, bank guarantees, implementation services) may justify exceeding the MEP if applicable. The Partner shall retain records demonstrating compliance with these requirements upon request by IBM or Mindware."""),
        
        ("C71", """These Terms and Conditions are binding and form an integral part of any agreement between Mindware FZ-LLC and the Partner. 				
Mindware may include additional terms in its quotations before its acceptance by the Partner, which the Partner agrees to be bound by as if fully incorporated into the agreement.
Notes:				
It is understood that final coverage dates for New Licenses and Software Subscription and Support Reinstatement part numbers will be based upon IBM's acceptance of the Purchase Order, as defined by the Passport Advantage program, irrespective of the dates which may appear in this special bid.				
Pricing is valid for this transaction only.

For Future Offers, any margin discounts, incentives, or promotions will be applied based on then current availability and terms at the actual time of billing for the Future Offer. For clarity, this means that the pricing for any Future Offering in this Quotation will be recalculated at the time of Distributor's actual order to apply existing applicable margin discounts, incentives, and promotions in effect at that time.				
				
This quote is contingent upon all paperwork being on file with IBM.				
				
Thank you for returning this quotation duly signed and stamped for your approval.


Signature                                                 ……………………………………….

Name                                                      ………………………………………

Title                                                      ………………………………………

Date                                                       ………………………………………""", {"merge_rows": 5})
    
    ]
