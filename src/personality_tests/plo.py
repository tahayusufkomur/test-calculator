import pandas as pd

from src.base_personality_test import BasePersonalityTest


class PLO(BasePersonalityTest):

    name = "PLO"
    type = 3
    description = "Potansiyel liderlik ölçeği"

    scoring = {
        'Hiç Uygun Değil': 1,
        'Çok Az Uygun': 2,
        'Biraz Uygun': 3,
        'Çoğunlukla Uygun': 4,
        'Tamamen Uygun': 5
    }

    categories_dict = {
        1: ["TOPLAM", [i for i in range(1, 16)]]
    }

    def format_report(self):

        df = pd.read_excel(self.output_path, engine='openpyxl')
        total_mean = df['TOPLAM'].mean()
        total_t_mean = df['TOPLAM T'].mean()

        df = df.style \
            .applymap(self.adjust_font) \
            .applymap(lambda x: self.mean_highlighter(x, total_mean), subset=pd.IndexSlice[:, ['TOPLAM']]) \
            .applymap(lambda x: self.mean_highlighter(x, total_t_mean), subset=pd.IndexSlice[:, ['TOPLAM T']])

        df.to_excel(self.output_path, engine='openpyxl', index=False)
        self.add_barchart("Sheet1", 'TOPLAM', 'TOPLAM', "İsim Soyisim", self.output_path)


if __name__ == '__main__':
    plo = PLO("/tests/resources/plo/PLO.xlsx", None)
    plo.create_report()

