import calendar
import locale
import string

import emoji

from Config import Config


class Post:
    def __init__(self, id="", enlace="", mensaje="", tipo="", fecha="", personas_alcanzadas=0, cant_coment=0,
                 cant_share=0, cant_likes=0, engagement=0, matching_audience=0, cant_feed_neg=0):

        self.cnf = Config()

        # arreglo las variables que aparecen como strings vacíos
        if cant_coment == "":
            cant_coment = 0

        if cant_share == "":
            cant_share = 0

        if cant_likes == "":
            cant_likes = 0

        self.id = id
        self.enlace = enlace
        self.mensaje = mensaje
        self.tipo = tipo
        self.fecha = fecha
        self.personas_alcanzadas = personas_alcanzadas
        self.cant_coment = cant_coment
        self.cant_share = cant_share
        self.cant_likes = cant_likes
        self.engagement = engagement
        self.matching_audience = matching_audience
        self.feed_back_negativo = cant_feed_neg

        # suma ponderada de los atributos
        self.indice_interacciones = self.get_interacciones()
        self.indice_completo = self.get_indice_mas_completo()

    def get_interacciones(self):
        return round(
            (self.cant_likes * self.cnf.LIKE_WEIGHT) + (self.cant_coment * self.cnf.COMMENT_WEIGHT) +
            (self.cant_share * self.cnf.SHARED_WEIGHT), 2
        )

    def get_indice_mas_completo(self):
        return round(
            (self.indice_interacciones * self.cnf.INTERACCIONES_WEIGHT) +
            (self.matching_audience * self.cnf.MATCH_AUDIENCE_WEIGHT) +
            (self.engagement * self.cnf.ENGAGEMENT_WEIGHT) +
            self.personas_alcanzadas * self.cnf.ALCANCE_WEIGHT,
            2
        )

    def text_has_emoji(self):
        return bool(emoji.get_emoji_regexp().search(self.mensaje))

    def text_has_hashtag(self):
        return self.mensaje.find("#") != -1

    def find_horario(self):
        horario = ""
        result_split_by_space = str(self.fecha).split()
        hora = int(result_split_by_space[1].split(":")[0])
        if 0 <= hora <= 5:
            horario = "madrugada"
        elif 6 <= hora < 12:
            horario = "mañana"
        elif 12 <= hora <= 18:
            horario = "tarde"
        elif 18 < hora <= 23:
            horario = "noche"
        return horario

    def get_week_day(self):
        # idioma en función del sistema operativo
        locale.setlocale(locale.LC_TIME, '')
        return calendar.day_name[self.fecha.weekday()]

    # no cuenta los emojis ni los caracteres solos Ej. ? -> no lo cuenta como palabra
    def count_words(self):
        return sum([i.strip(string.punctuation).isalpha() for i in self.mensaje.split()])

    def generar_preview(self, tam_mensaje=8):
        mensaje = str(self.mensaje)
        split = mensaje.split(" ")
        tam = len(split)
        if tam > tam_mensaje:
            tam = tam_mensaje
        return " ".join(split[:tam]) + "..."

    def get_dia_mes_ano(self):
        return self.fecha.strftime('%Y-%m-%d')
