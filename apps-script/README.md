# Comptes élèves du Banc de câblage

Le script `Code.gs` relie le Banc de câblage à un Google Sheet : connexion des élèves, accès autorisés par le professeur, suivi des exercices.

## Mise en route (une seule fois, avec le compte de l'école)

1. Créer un Google Sheet vide nommé **Banc de câblage – données**.
2. **Extensions › Apps Script**. Effacer le contenu de `Code.gs`, coller tout le fichier [`Code.gs`](Code.gs), enregistrer.
3. En haut, choisir la fonction **installer** puis **Exécuter**. Accepter les autorisations (Google affiche « application non validée » : c'est votre propre script, cliquer sur *Paramètres avancés › Accéder à…*). Le script crée les onglets Profs, Classes, Autorisés, Demandes, Résultats et Journal.
4. Créer les deux déploiements : **Déployer › Nouveau déploiement › roue dentée › Application Web**.

   | Déploiement | Description | Exécuter en tant que | Qui a accès |
   |---|---|---|---|
   | A | Connexion | Moi | Tous les utilisateurs de elmarche.be |
   | B | Données | Moi | Tout le monde |

5. Copier l'**URL de l'application Web** de chacun (elle se termine par `/exec`) et les mettre dans l'app (`var CPT = {loginUrl: A, apiUrl: B}` dans `index.html`).

## Onglets

- **Profs** : adresses des professeurs (accès complet et page « Mes élèves »).
- **Classes** : une ligne par classe, avec l'adresse du professeur.
- **Autorisés** : adresses autorisées, ou tout un domaine en écrivant `@domaine.be`.
- **Demandes** : premières connexions en attente d'accord. « Identité vérifiée » = l'adresse a été confirmée (compte Google ou code par e-mail).
- **Résultats** : une ligne par élève et par exercice (essais, réussite, meilleure note de recherche de pannes).
- **Journal** : chaque essai, avec sa date.

## Facultatif

Dans **Paramètres du projet › Propriétés du script**, ajouter `CODE_PROF` = votre code de licence prof : les professeurs connectés passent alors automatiquement en mode prof.

## Mettre à jour le script

Recoller le code, exécuter **installer**, puis **Gérer les déploiements › crayon › Nouvelle version**, pour chacun des deux déploiements (les adresses ne changent pas).
