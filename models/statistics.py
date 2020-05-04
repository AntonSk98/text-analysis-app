class Statistics:
    def __init__(self, text_number, post_number, id_numbers):
        self.text_number = text_number
        self.post_number = post_number
        self.id_numbers = id_numbers

    def __str__(self):
        return f'{self.text_number} {self.post_number}'
