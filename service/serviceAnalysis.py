import glob
import os
import re
from collections import defaultdict
import xlsxwriter
import matplotlib.pyplot as plt
import numpy
import openpyxl
import pandas as pd
from nltk.tokenize import word_tokenize
from openpyxl import load_workbook

from models.forum_row import ForumRow
from models.statistics import Statistics

merged_document_full_path = ''
all_words = {}
forum_data = []
counter_by_words = {}
counter_by_post = {}
statistics = {}
posts_count = 0
words_count = 0


def merge_csv_files(path_to_csv_files):
    global merged_document_full_path
    merged_document_full_path = path_to_csv_files + '\\merged_document.csv'
    if os.path.isfile(merged_document_full_path):
        os.remove(merged_document_full_path)
    os.chdir(path_to_csv_files)
    extension = 'csv'
    all_filenames = [i for i in glob.glob('*.{}'.format(extension))]
    combined_csv = pd.concat([pd.read_csv(f) for f in all_filenames])
    combined_csv.to_csv("merged_document.csv", index=False, encoding='utf-8-sig')



def are_entries_empty(entries, app):
    for name_of_entry in entries:
        if entries[name_of_entry] == "":
            app.errorBox(title='Error!', message='Please fill in all fields!')
            return True
    else:
        return False


def get_dictionary_of_words(path_to_dictionary):
    df = pd.ExcelFile(path_to_dictionary)
    sheet_names = df.sheet_names
    global all_words
    all_words = {}
    for sheet_name in sheet_names:
        df = pd.read_excel(path_to_dictionary, header=None, index_col=False, sheet_name=sheet_name)
        all_words[sheet_name] = numpy.concatenate(df.values, axis=0)
        all_words[sheet_name] = all_words[sheet_name][~pd.isnull(all_words[sheet_name])]


def get_all_data_from_merged_document():
    global merged_document_full_path
    global forum_data
    forum_data = []
    csv_file = pd.read_csv(merged_document_full_path, usecols=['Id', 'Topic', 'Text'])
    ids = csv_file['Id']
    data_topics = csv_file['Topic']
    data_texts = csv_file['Text']
    for row_id, topic, text in zip(ids, data_topics, data_texts):
        forum_data.append(ForumRow(row_id, topic, text))


def clear_row(text):
    text = text.replace('.', ' ')
    emoji_pattern = re.compile("[" # turn it into a function
                               u"\U0001F600-\U0001F64F"
                               u"\U0001F300-\U0001F5FF"
                               u"\U0001F680-\U0001F6FF"
                               u"\U0001F1E0-\U0001F1FF"
                               u"\U00002702-\U000027B0"
                               u"\U000024C2-\U0001F251"
                               u"\U0001f926-\U0001f937"
                               u'\U00010000-\U0010ffff'
                               u"\u200d"
                               u"\u2640-\u2642"
                               u"\u2600-\u2B55"
                               u"\u23cf"
                               u"\u23e9"
                               u"\u231a"
                               u"\u3030"
                               u"\ufe0f"
                               "]+", flags=re.UNICODE)
    return emoji_pattern.sub(r' ', text)


def run_analyzer(key_word):
    global all_words, forum_data, statistics, posts_count, words_count
    statistics = defaultdict(list)
    lists = all_words
    for i in lists:
        if i not in statistics:
            statistics[i] = {}
    posts_count = 0
    words_count = 0
    for rowObject in forum_data:
        list_words = []
        if key_word == 'texts':
            text = word_tokenize(clear_row(str(rowObject.text)))
        elif key_word == 'posts':
            text = word_tokenize(clear_row(str(rowObject.topic)))
        else:
            text = word_tokenize(clear_row(str(rowObject.topic)+' '+str(rowObject.text)))
        posts_count += 1
        words_count += len(text)
        for word in text:
            lower_word = word.lower()
            for key in lists:
                for value in lists[key]:
                    if lower_word == value.lower():
                        if lower_word in statistics[key]:
                            count = statistics[key][lower_word].text_number
                            if lower_word not in list_words:
                                post_count = statistics[key][lower_word].post_number + 1
                            else:
                                post_count = statistics[key][lower_word].post_number
                        else:
                            count = 0
                            post_count = 1
                        statistics[key][lower_word] = Statistics(count+1, post_count, rowObject.rowId)
                        list_words.append(lower_word)
    global counter_by_words, counter_by_post
    counter_by_words = defaultdict(list)
    counter_by_post = defaultdict(list)
    for criteria in statistics:
        list_ids = []
        count_for_text = 0
        count_for_posts = 0
        if criteria not in counter_by_words:
            counter_by_words[criteria] = 0
        if criteria not in counter_by_post:
            counter_by_post[criteria] = 0
        for word in statistics[criteria]:
            if statistics[criteria][word].id_number not in list_ids:
                count_for_posts += statistics[criteria][word].post_number
            count_for_text += statistics[criteria][word].text_number
            list_ids.append(statistics[criteria][word].id_number)
        counter_by_words[criteria] = count_for_text
        counter_by_post[criteria] = count_for_posts


def save_result_file(path, file_name):
    global statistics, counter_by_words, counter_by_post, posts_count, words_count
    os.chdir(path)
    writer = pd.ExcelWriter(file_name+'.xlsx', engine='xlsxwriter')
    for criteria in statistics:
        words = []
        number_words = []
        number_posts = []
        data = list(statistics[criteria].items())
        an_array = numpy.array(data)
        for element in an_array:
            words.append(element[0])
            number_words.append(int(element[1].text_number))
            number_posts.append(int(element[1].post_number))
        pd.DataFrame({'Word': words,
                      'Words\' number': number_words,
                      'Posts\' number': number_posts}).to_excel(writer, sheet_name=criteria, index=False)
    words_counter = list(counter_by_words.items())
    words_counter_array = numpy.array(words_counter)
    word_type = []
    number_type = []
    for array in words_counter_array:
        word_type.append(array[0])
        number_type.append(int(array[1]))
    pd.DataFrame({'Type': word_type,
                  'Words\' number': number_type}).to_excel(writer, sheet_name='StatisticsByWords', index=False)
    posts_counter = list(counter_by_post.items())
    posts_counter_array = numpy.array(posts_counter)
    word_type = []
    number_type = []
    for array in posts_counter_array:
        word_type.append(array[0])
        number_type.append(int(array[1]))
    pd.DataFrame({'Type': word_type,
                  'Posts\' number': number_type}).to_excel(writer, sheet_name='StatisticsByPosts', index=False)
    pd.DataFrame({'Number of posts': [posts_count],
                  'Number of words': [words_count]}).to_excel(writer, sheet_name='GeneralInformation', index=False)
    writer.save()
    writer.close()
    save_picture(path+'/'+file_name+'.xlsx', 'StatisticsByWords', 'Words\' number', 'Statistics by words', file_name)
    save_picture(path+'/'+file_name+'.xlsx', 'StatisticsByPosts', 'Posts\' number', 'Statistics by posts',file_name)


def save_picture(file_path, sheet_name, second_column_name, title, analysis_type):
    df = pd.read_excel(file_path, sheet_name=sheet_name, index=False)
    type_word = df["Type"]
    number = df[second_column_name]
    label = []
    sum_number = sum(number)
    if sum_number == 0:
        return
    for classification_type, count in zip(type_word, number):
        label.append(classification_type+':'+str(round(count/sum_number, 2) * 100)+'%')
    plt.pie(number, labels=label, shadow=True, startangle=140)
    plt.axis('equal')
    plt.tight_layout()
    figure = plt.gcf()
    figure.set_size_inches(8, 6)
    lgd = plt.legend(title=title,
                     loc="best",
                     bbox_to_anchor=(1, 0, 0.5, 1))

    plt.savefig(analysis_type+'#'+sheet_name+'.png', bbox_extra_artists=(lgd,), bbox_inches='tight', dpi=70)
    plt.close(figure)
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    img = openpyxl.drawing.image.Image(analysis_type+'#'+sheet_name+'.png')
    img.anchor = 'D1'
    ws.add_image(img)
    wb.save(file_path)
