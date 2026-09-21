# AO Manager

Application de gestion du cycle de vie des appels d'offres (cadrage, documents, exigences,
critères, fournisseurs, évaluation, comparaison, synthèse), avec base de données partagée et
comptes utilisateurs — chaque personne ne voit que ses propres AO.

## Stack

- **Frontend** : React + Vite
- **Base de données + authentification** : Supabase (Postgres + Auth, palier gratuit)
- **Hébergement** : Netlify (palier gratuit)

## Développement local

```bash
npm install
npm run dev
```

L'app se connecte à un projet Supabase déjà créé et configuré (voir `src/supabaseClient.js` —
la clé qui s'y trouve est la clé publique "anon", normale à exposer côté client ; la vraie
sécurité vient des règles Row Level Security côté base de données, pas de cette clé).

## Base de données

Le projet Supabase (`ao-manager`, région eu-central-1) contient une table `tenders` :

- `id` (uuid, clé primaire)
- `user_id` (référence vers l'utilisateur propriétaire)
- `reference`, `title` (colonnes pratiques pour lister/filtrer)
- `data` (jsonb — l'intégralité du dossier AO)
- `created_at`, `updated_at`

La sécurité par ligne (RLS) est activée : un utilisateur ne peut lire, créer, modifier ou
supprimer que les lignes où `user_id` correspond à son propre compte. C'est appliqué côté
base de données, donc impossible à contourner depuis le frontend.

## Comptes utilisateurs

L'authentification se fait par email + mot de passe (Supabase Auth). Par défaut, Supabase
envoie un email de confirmation à la création d'un compte — la personne doit cliquer sur le
lien reçu avant de pouvoir se connecter. Pour désactiver cette confirmation (utile pour une
petite équipe interne) : dans le tableau de bord Supabase → Authentication → Providers → Email
→ décocher "Confirm email".

## Pousser le projet sur GitLab

Depuis ce dossier, sur votre machine :

```bash
git init
git add .
git commit -m "Initial commit — AO Manager"
git branch -M main
git remote add origin <URL_DE_VOTRE_DEPOT_GITLAB>
git push -u origin main
```

(Remplacez `<URL_DE_VOTRE_DEPOT_GITLAB>` par l'URL de votre dépôt GitLab, créé au préalable
sur gitlab.com ou votre instance interne.)

## Déploiement (Netlify)

Si vous connectez ensuite le dépôt GitLab à Netlify (Netlify → Add new site → Import from Git),
Netlify lira automatiquement `netlify.toml` (commande `npm run build`, dossier `dist`) et
redéploiera à chaque push. Pensez à renseigner en variables d'environnement Netlify les mêmes
valeurs Supabase que dans `src/supabaseClient.js` si vous préférez ne pas les committer en dur.
