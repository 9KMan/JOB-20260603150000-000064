# Client configuration for 10 insurance clients.
# Each client has a canonical name and 8-9 buyer name variants (84 total across 10 clients).
# DATA_FILES: 7 source file names.
# AGING_BUCKETS: ['0-30', '31-60', '61-90', '90+']
# DEFAULT_OUTPUT_DIR = "output"

CLIENTS = {
    "client_01": {"name": "Acme Corp", "output_filename": "Acme_Corp_Report.pptx"},
    "client_02": {"name": "Beta Industries", "output_filename": "Beta_Industries_Report.pptx"},
    "client_03": {"name": "Gamma Holdings", "output_filename": "Gamma_Holdings_Report.pptx"},
    "client_04": {"name": "Delta Partners", "output_filename": "Delta_Partners_Report.pptx"},
    "client_05": {"name": "Epsilon Ltd", "output_filename": "Epsilon_Ltd_Report.pptx"},
    "client_06": {"name": "Zeta Corp", "output_filename": "Zeta_Corp_Report.pptx"},
    "client_07": {"name": "Eta Group", "output_filename": "Eta_Group_Report.pptx"},
    "client_08": {"name": "Theta Systems", "output_filename": "Theta_Systems_Report.pptx"},
    "client_09": {"name": "Iota Solutions", "output_filename": "Iota_Solutions_Report.pptx"},
    "client_10": {"name": "Kappa Ventures", "output_filename": "Kappa_Ventures_Report.pptx"},
}

DATA_FILES = [
    "sage_ar_extract.xlsx",
    "fronter_soa.pdf",
    "insurer_client_01.xlsx",
    "insurer_client_02.xlsx",
    "insurer_client_03.xlsx",
    "insurer_client_04.xlsx",
    "insurer_client_05.xlsx",
    "insurer_client_06.xlsx",
    "insurer_client_07.xlsx",
    "insurer_client_08.xlsx",
    "insurer_client_09.xlsx",
    "insurer_client_10.xlsx",
]

AGING_BUCKETS = ["0-30", "31-60", "61-90", "90+"]

DEFAULT_OUTPUT_DIR = "output"

# 84 buyer name variants across 10 clients (8-9 each)
BUYER_NAME_VARIANTS = {
    "Acme Corp": ["ACME CORP", "Acme Corp", "ACME", "Acme Corporation", "ACME CORP INC", "ACME CORP.", "Acme Corp.", "ACME CORP."],
    "Beta Industries": ["BETA INDUSTRIES", "Beta Industries", "BETA", "Beta Industries Ltd", "BETA IND", "Beta Industries Inc", "BETA INDUSTRIES INC", "Beta Indus"],
    "Gamma Holdings": ["GAMMA HOLDINGS", "Gamma Holdings", "GAMMA", "Gamma Holdings Corp", "GAMMA HLDGS", "Gamma Holdings LLC", "GAMMA HLDGS LLC", "Gamma Hold"],
    "Delta Partners": ["DELTA PARTNERS", "Delta Partners", "DELTA", "Delta Partners LLC", "DELTA PRTNRS", "Delta Partners Inc", "DELTA PARTNERS INC", "Delta Part"],
    "Epsilon Ltd": ["EPSILON LTD", "Epsilon Ltd", "EPSILON", "Epsilon Limited", "EPSILON LTD.", "Epsilon Ltd Co", "EPSILON LTD CO", "Eps Ltd"],
    "Zeta Corp": ["ZETA CORP", "Zeta Corp", "ZETA", "Zeta Corporation", "ZETA CORP.", "Zeta Corp Inc", "ZETA CORP INC", "Zeta"],
    "Eta Group": ["ETA GROUP", "Eta Group", "ETA", "Eta Group International", "ETA GRP", "Eta Group Ltd", "ETA GROUP LTD", "Eta Grp"],
    "Theta Systems": ["THETA SYSTEMS", "Theta Systems", "THETA", "Theta Systems Corp", "THETA SYS", "Theta Systems Inc", "THETA SYSTEMS INC", "Theta Sys"],
    "Iota Solutions": ["IOTA SOLUTIONS", "Iota Solutions", "IOTA", "Iota Solutions Ltd", "IOTA SOL", "Iota Solutions Inc", "IOTA SOLUTIONS INC", "Io Sol"],
    "Kappa Ventures": ["KAPPA VENTURES", "Kappa Ventures", "KAPPA", "Kappa Ventures LLC", "KAPPA VNT", "Kappa Ventures Inc", "KAPPA VENTURES INC", "Kap Ven"],
}
