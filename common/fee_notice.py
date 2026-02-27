def _render_fee_notice_message(template_content, sender_name, sender_phone, mention, unpaid_detail, fee_sheet_url):
    """회비 고지 템플릿 플레이스홀더를 실제 값으로 치환."""
    message = template_content.replace('{발신자}', sender_name)
    message = message.replace('{전화번호}', sender_phone)
    message = message.replace('{멘션}', mention)
    message = message.replace('{미납내역}', unpaid_detail)
    message = message.replace('{납부문서URL}', fee_sheet_url)
    return message
