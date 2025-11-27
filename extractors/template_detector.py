# extractors/template_detector.py
import fitz


def detect_ibm_template(file_like) -> str:
    """
    Auto-detect IBM template based on structural differences
    Template 1: Parts Information with coverage dates
    Template 2: Software as a Service with subscription parts
    """
    try:
        doc = fitz.open(stream=file_like.read(), filetype="pdf")
        sample_text = ""
        for page_num in range(min(3, len(doc))):
            page = doc[page_num]
            sample_text += page.get_text("text") or page.get_text()
        doc.close()
        file_like.seek(0)
        
        text_lower = sample_text.lower()
        
        # TEMPLATE 1 INDICATORS: Parts-based structure
        has_parts_information = 'parts information' in text_lower
        has_coverage_dates = 'coverage start' in text_lower or 'coverage end' in text_lower
        has_entitled_svp = 'entitled unit svp' in text_lower or 'entitled ext svp' in text_lower
        has_disc_percentage = 'disc %' in text_lower and 'bid unit svp' in text_lower
        
        # TEMPLATE 2 INDICATORS: Service/Subscription-based structure
        has_saas_header = 'software as a service' in text_lower
        has_subscription_part = 'subscription part#' in text_lower or 'subscription part:' in text_lower
        has_service_level = 'service level agreement' in text_lower
        has_subscription_length = 'subscription length' in text_lower
        has_billing_type = 'billing: upfront' in text_lower or 'billing: annual' in text_lower
        has_commit_value = 'total commit value' in text_lower and 'customer entitled' in text_lower
        has_renewal_type = 'renewal type:' in text_lower
        
        # Score templates
        template1_score = sum([
            has_parts_information,
            has_coverage_dates,
            has_entitled_svp,
            has_disc_percentage
        ])
        
        template2_score = sum([
            has_saas_header,
            has_subscription_part,
            has_service_level,
            has_subscription_length,
            has_billing_type,
            has_commit_value,
            has_renewal_type
        ])
        
        # Decision with clear preference
        if template2_score >= 3:
            return 'template2'
        elif template1_score >= 2:
            return 'template1'
        else:
            # Final fallback checks
            if 'software as a service' in text_lower or 'subscription part' in text_lower:
                return 'template2'
            elif 'parts information' in text_lower or 'coverage start' in text_lower:
                return 'template1'
            else:
                return 'template1'  # Default
                
    except Exception as e:
        print(f"Detection error: {e}")
        return 'template1'