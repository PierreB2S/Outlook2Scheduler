"""Erreurs affichées telles quelles à l'utilisateur (pas de trace Python)."""


class ErreurUtilisateur(Exception):
    """Erreur attendue : le message suffit à comprendre quoi faire."""


class ErreurConfig(ErreurUtilisateur):
    pass


class ErreurOutlook(ErreurUtilisateur):
    pass


class ErreurOutlookFerme(ErreurOutlook):
    pass


class ErreurSecuriteOutlook(ErreurOutlook):
    pass


class ErreurReseau(ErreurUtilisateur):
    pass


class ErreurSSL(ErreurReseau):
    pass


class ErreurSondage(ErreurUtilisateur):
    pass


class ErreurFormatPage(ErreurSondage):
    pass


class ErreurSondageJoursSemaine(ErreurSondage):
    pass


class ErreurMotDePasse(ErreurSondage):
    pass
