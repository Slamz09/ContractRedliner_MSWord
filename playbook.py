# playbook.py
playbook = {
    "payment_customer": {
        "keywords": ["payment", "billing", "invoice", "compensation", "terms of payment"],
        "context_keywords": ["customer shall pay", "client shall pay", "net 30"],
        "preferred": "Payment terms shall be Net 30 days.",
    },
    "payment_vendor": {
        "keywords": ["payment", "billing", "invoice", "compensation", "terms of payment"],
        "context_keywords": ["vendor shall be paid", "clean harbors shall pay", "net 60"],
        "preferred": "Payment terms shall be Net 60 days.",
    },
    "termination": {
        "keywords": ["termination", "cancel", "end", "notice", "termination rights"],
        "preferred": "Either party may terminate with 30 days' written notice.",
    }
}