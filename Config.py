from utiles.Sigleton import SingletonMeta


class Config(metaclass=SingletonMeta):
    def __init__(self):
        self.TOP = 3

        # posicion de las columnas en el excel
        self.ID = 0
        self.ENLACE = 1
        self.MENSAJE = 2
        self.TIPO = 3
        self.FECHA_PUBLICACION = 6
        self.ALCANCE = 8
        self.ENGAGEMENT = 14
        self.MATCH_AUDIENCE = 15

        self.SHARED = 10
        self.LIKE = 11
        self.COMMENT = 12

        self.FEED_BACK_NEGATIVE = 17

        # peso para sacar índice interacciones
        self.SHARED_WEIGHT = 0.4
        self.LIKE_WEIGHT = 0.3
        self.COMMENT_WEIGHT = 0.3

        # peso para índice mejores post
        self.INTERACCIONES_WEIGHT = 0.4
        self.MATCH_AUDIENCE_WEIGHT = 0.3
        self.ENGAGEMENT_WEIGHT = 0.2
        self.ALCANCE_WEIGHT = 0.1

        self.show_alcance = True
        self.show_comments = True
        self.show_reacciones = True
        self.show_compartidos = True
        self.show_indice_interacciones = True
        self.show_engagement = True
        self.show_match_audience = True
        self.show_indice_completo = True
        self.show_feed_back_negativo = True
