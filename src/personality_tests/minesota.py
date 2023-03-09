import pandas as pd

from src.base_personality_test import BasePersonalityTest


class Minesota(BasePersonalityTest):

    name = "Minesota"
    type = 2
    description = "Minesota"

    scoring = {
        'Hiç Memnun değilim': 1,
        'Memnun değilim': 2,
        'Kararsızım': 3,
        'Memnunum': 4,
        'Çok Memnunum': 5
    }

    categories_dict = {1: ["İçsel Doyum", [1, 2, 3, 4, 7, 8, 9, 10, 11, 15, 19, 20]],
                       2: ['Dışsal Doyum', [5, 6, 12, 13, 14, 16, 17, 18]],
                       3: ['Genel Doyum', [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]]
                       }

    def format_report(self):

        column_1 = 'İçsel Doyum'
        column_2 = 'Dışsal Doyum'
        column_3 = 'Genel Doyum'

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
            .applymap(lambda x: self.mean_highlighter(x, mean_t_2), subset=pd.IndexSlice[:, [column_2+" T"]])            \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_3), subset=pd.IndexSlice[:, [column_3+" T"]]) \


        df.to_excel(self.output_path, engine='openpyxl', index=False)
        self.add_barchart("Sheet1", column_1, column_3, "İsim Soyisim", self.output_path)


if __name__ == '__main__':
    mine = Minesota("/Users/tkoemue/ws/projects/merve/five-factor-test-calculator/yanıtlar/MİTÖ.xlsx", None)
    mine.create_report()