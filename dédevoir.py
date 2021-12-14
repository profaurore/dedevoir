# https://github.com/vgrem/Office365-REST-Python-Client/tree/master/office365

import json
import os
import os.path
from datetime import datetime
import getpass
import logging
import traceback
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.runtime.http.request_options import RequestOptions
from office365.runtime.http.request_options import RequestOptions


ctx = None
uri_serveur = 'https://csviamonde.sharepoint.com/'
uri_site = None
utilisateur = None
répertoires_élèves = None
travail = None


def menu(items, invite):
    print()

    for i, item in enumerate(items):
        print(f'{i + 1}. {item}')

    print()

    while True:
        choix = input(invite).strip()
        if choix == 'q':
            return

        try:
            choix = int(choix) - 1
            if 0 <= choix < len(items):
                break
        except ValueError:
            pass

    return choix


def sélectionner_serveur():
    print(f'URI du serveur: {uri_serveur}')

    return sélectionner_connexion


def sélectionner_connexion():
    global ctx, uri_serveur, utilisateur

    while True:
        print()

        while True:
            util_id = input('Identifiant: ').strip()
            if util_id == 'q':
                return
            if util_id != '':
                break

        if '@' not in util_id:
            util_id += '@csviamonde.ca'

        while True:
            util_passe = getpass.getpass('Mot de passe: ').strip()
            if util_passe == 'q':
                return
            if util_passe != '':
                break

        try:
            _utilisateur = UserCredential(util_id, util_passe)
            _ctx = ClientContext(uri_serveur).with_credentials(_utilisateur)
            break
        except Exception as e:
            print(f'L\'identifiant ou le mot de passe est invalide.')
            print(e)

    utilisateur = _utilisateur
    ctx = _ctx

    return sélectionner_site


def sélectionner_site():
    global ctx, uri_site, utilisateur

    try:
        with open('dédevoir_sites_enregistrés.txt', 'r') as f:
            sites_sauvegardés = [s.rsplit('\t', 1) for s in f.read().split('\n') if s != '']
    except Exception as e:
        sites_sauvegardés = []

    liste_choix = ['Liste automatique', 'Recherche', 'URI du site']
    if len(sites_sauvegardés):
        liste_choix.append('Sites sauvegardés')

    while True:
        mode = menu(liste_choix, 'No de l\'option: ')
        if mode is None:
            return

        if mode == 2:
            while True:
                _uri_site = input('\nURI du site: ').strip()
                if _uri_site == 'q':
                    return
                if _uri_site != '':
                    break

            infixe = 'sharepoint.com/sites/'
            pos = _uri_site.find(infixe)
            if pos == -1:
                print(f'L\'URI du site est invalide.')
                continue

            _uri_site = _uri_site[0:pos + len(infixe)] + _uri_site[pos + len(infixe):].split('/', 2)[0]

            try:
                _ctx = ClientContext(_uri_site).with_credentials(utilisateur)
            except Exception as e:
                print(f'L\'URI du site est invalide.\n  {str(e)}')
                continue

            while True:
                choix = input('\nSauvegarder le site? (o ou n): ').strip()
                if choix == 'q':
                    return
                if choix in 'on':
                    break

            if choix == 'o':
                try:
                    _ctx.load(_ctx.web)
                    _ctx.execute_query()
                    titre = _ctx.web.properties['Title']
                except:
                    print(f'Incapable de charger le site.\n  {str(e)}')
                    continue

                try:
                    with open('dédevoir_sites_enregistrés.txt', 'a') as f:
                        f.write(f'{titre}\t{_uri_site}\n')
                except Exception as e:
                    print(f'Incapable de sauvegarder le site.\n  {str(e)}')
                    continue

            break
        elif mode == 3:
            site = menu([f'{nom}\n    {uri}' for nom, uri in sites_sauvegardés], 'No du site: ')
            if site is None:
                return

            _uri_site = sites_sauvegardés[site][1]
            try:
                _ctx = ClientContext(_uri_site).with_credentials(utilisateur)
                break
            except Exception as e:
                print(f'L\'URI du site est invalide.\n  {str(e)}')
        else:
            recherche = input('\nRecherche: ').lower() if mode == 1 else ''
            section = 'SECTION_ETBR' if mode == 0 else ''

            try:
                données = ctx.execute_request_direct(RequestOptions(f'{uri_serveur}/_api/search/query?querytext=\'contentclass:STS_Site Path:"{uri_serveur}/sites/{section}*"\'&rowlimit=500&selectproperties=\'Title,Path\''))
            except Exception as e:
                print(f'La connexion au serveur a échouée.\n  {str(e)}')
                continue

            try:
                données_json = json.loads(données.content)
            except json.JSONDecodeError as e:
                print(f'Les données JSON reçues du serveur sont invalides.\n  {str(e)}')
                continue

            try:
                données_sites = données_json['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results']

                sites = []
                for données_site in données_sites:
                    site = {}
                    for données_item in données_site['Cells']['results']:
                        site[données_item['Key']] = données_item['Value']
                    if recherche in site['Title'].lower():
                        sites.append(site)
            except (KeyError, TypeError) as e:
                print(f'La structure des données JSON reçues du serveur est inattendue.\n  {str(e)}')
                continue

            if not len(sites):
                print('\nAucune équipe trouvée.')
                continue

            site = menu([f'{s["Title"]}\n    {s["Path"]}' for s in sites], 'No du site: ')
            if site is None:
                return

            _uri_site = sites[site]['Path']
            try:
                _ctx = ClientContext(_uri_site).with_credentials(utilisateur)
                break
            except Exception as e:
                print(f'L\'URI du site est invalide.\n  {str(e)}')

    ctx = _ctx
    uri_site = _uri_site

    return sélectionner_devoir


def sélectionner_devoir():
    global ctx, répertoires_élèves, travail

    try:
        répertoire = ctx.web.get_folder_by_server_relative_url('Travaux des tudiants/Fichiers envoyés')
        ctx.load(répertoire)
        ctx.execute_query()

        _répertoires_élèves = répertoire.folders
        ctx.load(_répertoires_élèves)
        ctx.execute_query()

        travaux = set()
        for répertoire_élève in _répertoires_élèves:
            répertoires_travaux = répertoire_élève.folders
            ctx.load(répertoires_travaux)
            ctx.execute_query()

            for répertoire_travail in répertoires_travaux:
                travaux.add(répertoire_travail.properties['Name'])
        travaux = list(travaux)
        travaux.sort()
    except Exception as e:
        print(f'\nLa structure du site est invalide.\n  {str(e)}')
        return sélectionner_site

    if not len(travaux):
        print('\nAucun travail trouvé.')
        return sélectionner_site

    idx_travail = menu(travaux, 'No du travail: ')
    if idx_travail is None:
        return
    _travail = travaux[idx_travail]

    travail = _travail
    répertoires_élèves = _répertoires_élèves

    return télécharger_devoirs


def télécharger_devoirs():
    global répertoires_élèves, travail, uri_serveur

    répertoire = datetime.now().strftime("%Y-%m-%d-%H%M%S - ") + travail
    répertoire_complet = os.path.abspath(répertoire)

    try:
        os.mkdir(répertoire)
    except FileExistsError:
        pass
    except OSError as e:
        print(f'Incapable de créer le répertoire « {répertoire_complet} ».\n  {str(e)}')
        return

    for répertoire_élève in répertoires_élèves:
        répertoire_travail = répertoire_élève.folders.filter(f'Name eq \'{travail}\'')
        ctx.load(répertoire_travail)
        ctx.execute_query()

        if len(répertoire_travail) == 0:
            continue

        élève = répertoire_élève.properties['Name']
        print(élève)

        télécharger_fichiers(répertoire_travail[0], répertoire, répertoire_complet, élève, 'v0')

        répertoires_versions = répertoire_travail[0].folders
        ctx.load(répertoires_versions)
        ctx.execute_query()

        for répertoire_version in répertoires_versions:
            version = répertoire_version.properties['Name'].replace('Version ', 'v')

            télécharger_fichiers(répertoire_version, répertoire, répertoire_complet, élève, version)

    print(f'\nLes fichiers ont étés sauvegardés dans le répertoire « {répertoire_complet} ».')


def télécharger_fichiers(répertoire_source, répertoire_cible, répertoire_cible_complet, élève, version):
    fichiers = répertoire_source.files
    ctx.load(fichiers)
    ctx.execute_query()

    for fichier in fichiers:
        print(f'  {version} {fichier.properties["Name"]}')

        try:
            données = fichier.read()
        except Exception as e:
            print(f'    Incapable de lire le fichier « {uri_serveur}/{fichier.properties["ServerRelativeUrl"]} ».\n      {str(e)}')
            continue

        nom = f'{élève} - {version} - {fichier.properties["Name"]}'
        try:
            with open(f'{répertoire_cible}/{nom}', 'wb') as f:
                f.write(données)
        except Exception as e:
            print(f'    Incapable d\'écrire le fichier « {répertoire_cible_complet}/{nom} ».\n      {str(e)}')
            continue


logging.disable(logging.CRITICAL)

try:
    étape = sélectionner_serveur
    while étape is not None:
        étape = étape()
except Exception as e:
    traceback.print_exc()
    print(e)

input('\nPesez la touche d\'entrée pour quitter...')
