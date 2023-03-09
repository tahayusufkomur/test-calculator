import pandas as pd
from src.base_personality_test import BasePersonalityTest, get_number_from_question


class FiveFactor(BasePersonalityTest):

    name = "b5kt"
    type = 1

    scoring = {
        'Hiç uygun değil': 1,
        'Uygun değil': 2,
        'Orta/kararsız': 3,
        'Biraz uygun': 4,
        'Çok uygun': 5
    }

    reverse_scoring = {
        'Hiç uygun değil': 5,
        'Uygun değil': 4,
        'Orta/kararsız': 3,
        'Biraz uygun': 2,
        'Çok uygun': 1
    }

    text_scoring = {
        5: '5',
        4: '4',
        3: '3',
        2: '2',
        1: '1'
    }

    should_be_same = [
        [31, 46],
        [6, 21, 26],
        [2, 22, 32],
        [7, 37],
        [17, 42, 47],
        [8, 28],
        [18, 33, 43],
        [9, 19, 49],
        [29, 24, 44],
        [34, 39],
        [4, 14],
        [25, 5],
        [5, 4],
        [15, 45],
        [10, 20, 30]
    ]
    reverse_calculation_list = [
        2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24,
        26, 28, 29, 30, 32, 34, 36, 38, 39, 44, 46, 49
    ]

    categories_dict = {1: ["Dışa dönüklük", [1, 6, 11, 16, 21, 26, 31, 36, 41, 46]],
                       2: ["Uyumluluk / Geçimlilik", [2, 7, 12, 17, 22, 27, 32, 37, 42, 47]],
                       3: ['Öz Denetim / Sorumluluk', [3, 8, 13, 18, 23, 28, 33, 38, 43, 48]],
                       4: ['Duygusal Denge', [4, 9, 14, 19, 24, 29, 34, 39, 44, 49]],
                       5: ['Gelişime Açıklık / Hayal Gücü', [5, 10, 15, 20, 25, 30, 35, 40, 45, 50]],
                       }

    def map_answers(self, df=None):

        for key, row in self.categories_dict.items():
            for column in self.test_questions:
                if get_number_from_question(column) in row[1]:
                    if get_number_from_question(column) in self.reverse_calculation_list:
                        self.test_df[column] = self.test_df[column].map(self.reverse_scoring)
                    else:
                        self.test_df[column] = self.test_df[column].map(self.scoring)

    def format_report(self):

        column_1 = 'Dışa dönüklük'
        column_2 = 'Uyumluluk / Geçimlilik'
        column_3 = 'Öz Denetim / Sorumluluk'
        column_4 = 'Duygusal Denge'
        column_5 = 'Gelişime Açıklık / Hayal Gücü'
        column_6 = 'Güven Endeksi'

        df = pd.read_excel(self.output_path, engine='openpyxl')
        mean_1 = df[column_1].mean()
        mean_2 = df[column_2].mean()
        mean_3 = df[column_3].mean()
        mean_4 = df[column_4].mean()
        mean_5 = df[column_5].mean()
        mean_6 = df[column_6].mean()
        mean_t_1 = df[column_1+" T"].mean()
        mean_t_2 = df[column_2+" T"].mean()
        mean_t_3 = df[column_3+" T"].mean()
        mean_t_4 = df[column_4+" T"].mean()
        mean_t_5 = df[column_5+" T"].mean()

        df = df.style \
            .applymap(self.adjust_font) \
            .applymap(lambda x: self.mean_highlighter(x, mean_1), subset=pd.IndexSlice[:, [column_1]])\
            .applymap(lambda x: self.mean_highlighter(x, mean_2), subset=pd.IndexSlice[:, [column_2]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_3), subset=pd.IndexSlice[:, [column_3]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_4), subset=pd.IndexSlice[:, [column_4]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_5), subset=pd.IndexSlice[:, [column_5]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_6), subset=pd.IndexSlice[:, [column_6]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_1), subset=pd.IndexSlice[:, [column_1+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_2), subset=pd.IndexSlice[:, [column_2+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_3), subset=pd.IndexSlice[:, [column_3+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_4), subset=pd.IndexSlice[:, [column_4+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_5), subset=pd.IndexSlice[:, [column_5+" T"]])

        df.to_excel(self.output_path, engine='openpyxl', index=False)

        self.add_barchart("Sheet1", column_1, column_5, "İsim Soyisim", self.output_path)


if __name__ == '__main__':
    x = FiveFactor("/Users/tkoemue/ws/projects/merve/five-factor-test-calculator/yanıtlar/B5KT.xlsx", None)
    x.create_report()
