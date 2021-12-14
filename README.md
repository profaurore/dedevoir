# Dédevoir

Ce logiciel télécharge tous les documents soumis à un devoir par les membres d'une équipe Teams. Une interface de ligne de commande est utilisée.

## Dépendances

Le paquet Python [Office365-REST-Python-Client](https://github.com/vgrem/Office365-REST-Python-Client) est requis pour accéder aux documents.
```
pip install Office365-REST-Python-Client
```

## Exécuter

```
python dédevoir.py
```

## Bâtir avec PyInstaller

```
pyinstaller --onefile --add-data "SAML.xml;office365/runtime/auth/providers/templates" dédevoir.py
```