import pandas as pd

from src.base_personality_test import BasePersonalityTest


class InclusiveLeadership(BasePersonalityTest):

    name = "Kapsayıcı Liderlik"
    type = 2
    description = "Kapsayıcı Liderlik"

    scoring = {
        'Hiç Katılmıyorum': 1,
        'Katılmıyorum': 2,
        'Kararsızım': 3,
        'Katılıyorum': 4,
        'Tamamen Katılıyorum': 5
    }

    categories_dict = {1: ["Açıklık", [1, 2, 3]],
                       2: ['Hazır Olma', [4, 5, 6, 7]],
                       3: ['Ulaşılabilirlik', [8, 9]]
                       }

    def format_report(self):

        column_1 = 'Açıklık'
        column_2 = 'Hazır Olma'
        column_3 = 'Ulaşılabilirlik'

        df = pd.read_excel(self.output_path, engine='openpyxl')
        mean_1 = df[column_1].mean()
        mean_2 = df[column_2].mean()
        mean_3 = df[column_3].mean()

        mean_t_1 = df[column_1+" T"].mean()
        mean_t_2 = df[column_2+" T"].mean()
        mean_t_3 = df[column_3+" T"].mean()

        df = df.style \
            .applymap(self.adjust_font) \
            .applymap(lambda x: self.mean_highlighter(x, mean_1), subset=pd.IndexSlice[:, [column_1]])\
            .applymap(lambda x: self.mean_highlighter(x, mean_2), subset=pd.IndexSlice[:, [column_2]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_3), subset=pd.IndexSlice[:, [column_3]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_1), subset=pd.IndexSlice[:, [column_1+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_2), subset=pd.IndexSlice[:, [column_2+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_3), subset=pd.IndexSlice[:, [column_3+" T"]]) \


        df.to_excel(self.output_path, engine='openpyxl', index=False)
        self.add_barchart("Sheet1", column_1, column_3, "Değerlendirdiğiniz kişiyi seçiniz.", self.output_path)


if __name__ == '__main__':
    x = InclusiveLeadership("empty.xlsx", None)
    x.create_report()