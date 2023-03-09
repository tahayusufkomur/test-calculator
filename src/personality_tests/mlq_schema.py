import json
import pathlib

import pandas as pd

from src.personality_tests.mlq_ast import MlqAst


class MlqSchema(MlqAst):

    name = "mlq-schema"
    current_path = pathlib.Path(__file__).parent.resolve()
    type = 2
    description = "MLQ Schema report"
    schema = json.load(
        open(f'{current_path}/files/schema.json', encoding="utf-8"))
    schema_report = {}
    schema_output_path = "raporlar/excel_reports/mlq_reports/schema_report.xlsx"

    def read_test_answers(self):
        ast_df = pd.read_excel(self.non_manager_path, engine='openpyxl')
        ust_df = pd.read_excel(self.manager_path, engine='openpyxl')

        ast_df.rename(columns={'Değerlendirdiğim kişiden daha alt konumdayım': 'Konum',
                               'Şifreyi giriniz ': 'Şifre',
                               'Değerlendirdiğiniz kişinin adını-soyadını ve  ünvanını eksiksiz  olacak şekilde yazınız. ': 'İsim Soyisim'}, inplace=True)
        ust_df.rename(columns={'Kendimi değerlendiriyorum.': 'Konum',
                               'İsim -Soyisim ve ünvanınızı eksiksiz olarak yazınız. ': 'İsim Soyisim'}, inplace=True)

        ast_df['Konum'] = ast_df['Konum'].map({'EVET': 'Ast'})
        ust_df['Konum'] = ust_df['Konum'].map({'EVET': 'Üst'})

        ust_df.columns = ast_df.columns
        df = ast_df.append(ust_df, ignore_index=True)
        self.managers = df[df['Konum'] == 'Üst']['İsim Soyisim'].tolist()
        return df

    def create_report(self):
        try:
            self.create_schema_report()
        except:
            print("Şema oluşturulamadı")

    def create_schema_report(self):
        report = {}
        output_path = self.schema_output_path

        for key, row in self.schema.items():
            df = self.test_df[self.test_df['İsim Soyisim'] == key].copy()
            df = df[df['Şifre'].str.contains('|'.join(self.passwords))]
            dones = df['Şifre'].tolist()
            done = []
            not_done = []
            # key in worker keys
            for ro in row:
                if ro in dones:
                    done.append(self.passwords[ro])
                else:
                    not_done.append(self.passwords[ro])

            try:
                ones_need_to_evaulate = ", ".join([self.passwords[ro] for ro in row])
            except:
                ones_need_to_evaulate = "Şifre Dosyası Eksik"

            report[key] = {'değerlendirenler': "Kimse Değerlendirmedi" if not done else ", ".join(done),
                           'değerlendirmeyenler': "Herkes Değerlendirdi" if not not_done else ", ".join(not_done),
                           'değerlendirmesi gerekenler': ones_need_to_evaulate}
        df = pd.DataFrame.from_dict(report, orient='index')
        df.to_excel(output_path, engine='openpyxl')

        # highlight değerlendirmeyenler
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        # page 1
        sheet_name = "Sheet1"
        df = pd.read_excel(output_path, engine='openpyxl', sheet_name=sheet_name)
        writer = pd.ExcelWriter(output_path, engine='openpyxl', mode='a')
        df = df.style \
            .applymap(self.adjust_font) \
            .applymap(self.negative_highlighter, subset=pd.IndexSlice[:, ['değerlendirmeyenler']])

        writer.book = wb
        wb.remove(wb[sheet_name])
        df.to_excel(excel_writer=writer, sheet_name=sheet_name, index=False)
        writer.close()

        self.adjust_width(output_path)


if __name__ == '__main__':
    x = MlqSchema(non_manager_path="/yanıtlar/MLQ-AST-2.xlsx",
                  manager_path="/Users/tkoemue/ws/projects/merve/five-factor-test-calculator/yanıtlar/MLQ-UST.xlsx",
                  colors=None)

    x.create_report()