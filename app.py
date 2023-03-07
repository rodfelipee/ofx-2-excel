from ofxparse import OfxParser
from csv import DictWriter
from glob import glob
import pandas as pd
import argparse

DATEFORMAT = "%m/%d/%Y"

class App():
    argparser = argparse.ArgumentParser()
    argparser.add_argument("-o", "--outputtype", help = "singlecsv, json, or manycsv", default="singlecsv")
    args = argparser.parse_args()

    outputtype = args.outputtype
    allStatements = []

    # bank_dict = {
    #     1: 'Banco do Brasil S.A.',
    #     33: 'Banco Santander (Brasil) S.A.',
    #     104: 'Caixa Econômica Federal',
    #     237: 'Banco Bradesco S.A.',
    #     341: 'Banco Itaú S.A.'
    # }

    def write_file(statement, out_file):

        with open(out_file, 'w', newline='') as f:
            fields = ['DATA', 'ID', 'TIPO', 'DESCRIÇÃO', 'VALOR', 'DEBITO', 'CREDITO', 'SALDO', 'CONTA', 'BANKID']
            wr = DictWriter(f, fieldnames=fields, delimiter=',', dialect='excel')
            wr.writeheader()
            for text in statement:
                wr.writerow(text)

    def get_stt_from_ofx(ofx):
        balance = ofx.account.statement.balance
        account = ofx.account
        # institution = ofx.account.institution

        statement = []
        credit_transactions = ['credit', 'dep', 'int', 'directdep']
        debit_transactions = ['debit', 'atm', 'pos',
                            'xfer', 'check', 'fee', 'payment']
        other_transactions = ['other']
        
        for transaction in ofx.account.statement.transactions:
            #print(transaction.type)
            credit = ""
            debit = ""
            balance = balance + transaction.amount
            if transaction.type in credit_transactions:
                credit = transaction.amount
            elif transaction.type in debit_transactions:
                debit = -transaction.amount
            elif transaction.type in other_transactions:
                if transaction.amount < 0:
                    debit = transaction.amount
                else:
                    credit = transaction.amount
            else:
                raise ValueError("Unknown transaction type:" + transaction.type)

            line = {
                'DATA': transaction.date.strftime(DATEFORMAT),
                'TIPO': transaction.type,
                'DESCRIÇÃO': transaction.memo,
                'ID': transaction.id,
                'VALOR': str(transaction.amount),
                'DEBITO': str(debit),
                'CREDITO': str(credit),
                'SALDO': str(balance),
                'CONTA': account.account_id,
                # 'EMPRESA': institution,
                'BANKID': account.routing_number,
                }
            statement.append(line)
        return statement
    
    def csv2xlsx(out_file):
        csv = pd.read_csv('ofx-transactions.csv', encoding='latin-1')
        excelWriter = pd.ExcelWriter('ofx-resume.xlsx')
        csv.to_excel(excelWriter, index_label='ABC', index=False, float_format='%2.f', freeze_panes=(1, 0))
        excelWriter.save()

    files = glob("*.ofx")
    for ofx_file in files:
        ofx = OfxParser.parse(open(ofx_file, encoding="latin-1"))

        statement = get_stt_from_ofx(ofx)
        allStatements = allStatements + statement

        if outputtype == 'manycsv':
            out_file = "converted_" + ofx_file.replace(".ofx", ".csv")
            write_file(statement, out_file)
        if outputtype == 'singlecsv':
            out_file = "ofx-transactions.csv"
            write_file(allStatements, out_file)
            csv2xlsx(out_file)

App()