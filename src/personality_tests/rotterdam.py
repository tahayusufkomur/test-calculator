import pandas as pd

from src.base_personality_test import BasePersonalityTest, get_number_from_question


class Rotterdam(BasePersonalityTest):

    name = "Rotterdam"
    type = 2

    scoring_a = {
        'a': 1,
        'b': 0,
    }
    scoring_b = {
        'a': 0,
        'b': 1
    }
    scoring_always_zero = {
        'a': 0,
        'b': 0
    }
    questions_a = [2, 6, 7, 9, 16, 17, 18, 20, 21, 23, 25, 29]
    questions_b = [3, 4, 5, 10, 11, 12, 13, 15, 22, 26, 28]

    categories_dict = {1: ["Şans ve Kadercilik", [18, 28, 2, 25, 13, 11, 9, 21]],
                       2: ['Siyasal Dış Kontrol', [12, 22, 17, 3, 29]],
                       3: ['Şans ve Kişiler Arası Dış Kontrol', [15, 16, 4]],
                       4: ['Okul Başarısında Dış Kontrol', [10, 5, 23]],
                       5: ['Kişilerarası ilişkilerde Dış Kontrol', [7, 26, 6, 20]],
                       6: ['Toplam', questions_a+questions_b],
                       }

    def map_answers(self):
        for column in self.test_questions:
            self.test_df[column] = self.test_df[column].apply(lambda x: str(x)[0])

        for column in self.test_questions:
            column_as_num = get_number_from_question(column)
            if column_as_num in self.questions_a:
                self.test_df[column] = self.test_df[column].map(self.scoring_a)
            elif column_as_num in self.questions_b:
                self.test_df[column] = self.test_df[column].map(self.scoring_b)
            else:
                self.test_df[column] = self.test_df[column].map(self.scoring_always_zero)

    def prepare_result(self):
        self.result_df = self.test_df.copy()[self.personal_questions]
        for key, row in self.categories_dict.items():
            self.result_df[row[0]] = self.sum_row(row[1])
            self.test_df[row[0]] = self.sum_row(row[1])
        # will be used for getting highest score categories
        self.text_calculation_df = self.test_df.copy()

    def format_report(self):

        column_1 = 'Şans ve Kadercilik'
        column_2 = 'Siyasal Dış Kontrol'
        column_3 = 'Şans ve Kişiler Arası Dış Kontrol'
        column_4 = 'Okul Başarısında Dış Kontrol'
        column_5 = 'Kişilerarası ilişkilerde Dış Kontrol'
        column_6 = 'Toplam'

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
        mean_t_6 = df[column_6+" T"].mean()

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
            .applymap(lambda x: self.mean_highlighter(x, mean_t_5), subset=pd.IndexSlice[:, [column_5+" T"]]) \
            .applymap(lambda x: self.mean_highlighter(x, mean_t_6), subset=pd.IndexSlice[:, [column_6+" T"]]) \
            .applymap(lambda x: self.character_highlighter(x), subset=pd.IndexSlice[:, ['Karakter']]) \

        df.to_excel(self.output_path, engine='openpyxl', index=False)
        self.add_barchart("Sheet1", column_1, column_6, "İsim Soyisim", self.output_path)

    def create_report(self):
        self.filter_df_by_password()
        if len(self.test_df) == 0:
            return 0
        # map answers to their corresponding score in dataframe
        self.map_answers()
        # remove questions from the result and add total column
        self.prepare_result()
        # add mean
        self.add_mean_row()
        # add standart deviation
        self.add_std_row()
        # add z-score
        self.add_tscore()
        # round df
        self.round_df()
        # add chacteristic
        self.add_characteristic()
        # add explanation
        self.add_text()
        # save
        self.save_report()

        self.copy_df = self.result_df.copy()
        # colorize
        self.format_report()
        # adjust width
        self.adjust_width()
        # add graphs sheet
        self.add_graphs_sheet()

        print(f"Created report for: {self.name}")

    def add_characteristic(self):
        self.result_df['Karakter'] = self.test_df.apply(map_characteristic, axis=1)

    def add_text(self):
        self.result_df['Açıklama'] = self.test_df.apply(map_explain, axis=1)

    def character_highlighter(self, x):
        if x == 'İç Kontrol Odaklı':
            return f"background-color: {self.colors['very_good']}; color: black;"
        elif x == 'Dış Kontrol Odaklı':
            return f"background-color: {self.colors['bad']}; color: black;"
        else:
            return None

    def rdo_highlighter(self, x):
        very_good = f"background-color: {self.colors['very_good'] if self.colors['very_good'] else 'white'}; color: black;"
        bad = f"background-color: {self.colors['bad'] if self.colors['very_good'] else 'white'}; color: black;"

        if x > 11:
            return bad
        if x <= 11:
            return very_good


def map_characteristic(row):
    if row['Toplam'] <= 11:
        return 'İç Kontrol Odaklı'
    if row['Toplam'] > 11:
        return 'Dış Kontrol Odaklı'


def map_explain(row):
    if row['Toplam'] <= 11:
        return "“İç kontrol odaklı” kişiler; davranışlarının ve kararlarının sorumluluklarını alırlar, yaşamlarının herhangi bir boyutuyla ilgili olarak mutsuz olduklarında, bunu kendi çabalarıyla değiştirebileceklerine inanırlar ve işi şansa, kadere bırakmak yerine çaba göstermek isterler."
    if row['Toplam'] > 11:
        return "“Dış kontrol odaklı” kişiler, davranışlarının ve karalarının sorumluluklarını almak istemezler. Davranışlarına ve kararlarına hep şans, kader, güçlü ve otoritesi olan diğer insanlar gibi kendileri dışındaki faktörlerin neden olduğunu düşünürler. Kontrol gücü kendisi dışında herhangi bir şey olabilir. Dış kontrol odaklı kişiler olumsuz olaylarda çaresizlik duygusuna sığınabilirler, olumlu olaylarda ise ödüllerin kendi çabalarından kaynaklanmadığına, yalnızca doğru zamanda doğru yerde olmanın getirdiği bir şans olduğuna inanırlar."


if __name__ == '__main__':
    rt = Rotterdam("/Users/tkoemue/ws/projects/merve/five-factor-test-calculator/yanıtlar/RDO.xlsx", None)
    rt.create_report()