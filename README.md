# Vulpix Pilotage — version Netlify (sans backend)

Même appli que la V1 artifact (MRR, TJM, CRM, temps, recommandations, export Excel),
hébergée sur ton propre Netlify. Aucune connexion externe : les données sont stockées
dans le navigateur (localStorage).

## Déploiement

Le site Netlify est relié à ce dépôt GitHub : chaque envoi (push) sur la branche `main`
déclenche automatiquement un nouveau déploiement. Plus besoin de zip ni de glisser-déposer.

Pour limiter la consommation de crédits Netlify, regrouper les modifications en un seul envoi.

## Tester en local

```bash
npm install
npm run dev
```
Ouvre `http://localhost:5173`.

## ⚠️ Limite importante à connaître

Les données sont stockées **dans le navigateur, sur cet appareil**. Concrètement :
- Si tu changes d'ordinateur ou de navigateur → tu ne les retrouves pas automatiquement
- Si tu vides le cache/les données de navigation → elles disparaissent
- Pas de synchronisation entre ton PC et ton téléphone

**Réflexe à prendre** : clique régulièrement sur **"Sauvegarde JSON"** (onglet Clients) pour
télécharger une copie de sécurité. Le bouton **"Restaurer"** te permet de réimporter ce
fichier n'importe quand — y compris pour transférer tes données d'un appareil à un autre.

## Pour aller plus loin (accès multi-appareils)

Si un jour tu veux consulter/modifier tes données depuis ton PC ET ton téléphone avec
synchronisation automatique, il faudra un vrai stockage partagé (Google Sheets, Airtable,
Supabase...). On avait exploré deux pistes :
- **Google Cloud + Sheets API** : plus robuste mais demande une carte bancaire pour activer
  certains services (même si l'usage réel reste gratuit)
- **Google Apps Script** : gratuit sans carte, mais plus délicat à déployer correctement
  (CORS, permissions d'accès)

Si tu veux repartir sur l'une de ces pistes plus tard, dis-le-moi.

## Structure du projet

```
src/App.jsx       → toute l'appli (dashboard, clients, temps, réglages)
src/lib/api.js     → lit/écrit dans localStorage
```

## Migrer tes données depuis la V1 artifact

1. Dans l'artifact Claude, clique **Sauvegarde JSON** (onglet Clients).
2. Une fois ce site déployé et ouvert, clique sur **Restaurer** (onglet Clients) et
   sélectionne ce fichier — toutes tes données (clients, temps, réglages) sont reprises.

## Seuils fiscaux (tableau de bord)

Le bloc « Seuils fiscaux » compare le CA estimé de l'année à trois repères : repère prudent
(29 000 €) et plafond haut estimé (33 000 €) du versement libératoire, et franchise de TVA
(37 500 €). Les seuils et le CA encaissé mois par mois se règlent dans Réglages. Sans saisie,
le CA est projeté à MRR constant. Les repères sont des estimations à valider avec un
expert-comptable.

## Ce qui reste identique à la V1

- Le calcul du TJM, MRR, net estimé, recommandations : code inchangé.
- L'export Excel : toujours disponible.
- Le design : strictement identique.

## Dépendance du MRR (tableau de bord)

Camembert « Dépendance du MRR » : part de chaque client actif dans le MRR, avec bascule « Par client / Par secteur ». Il se met à jour dès qu'un MRR, un statut ou un secteur change. Alerte sur le plus gros poste : vert < 20 %, jaune 20-25 %, orange 25-35 %, rouge > 35 %.

## Temps passé par client (tableau de bord)

Second camembert : part du temps saisi par client, par mois ou sur tous les mois. Une session partagée est répartie à parts égales (au prorata du nombre de posts pour les posts) ; les sessions sans client apparaissent en « Sans client précis ». Chaque ligne rappelle la part de MRR du client, et la phrase de synthèse signale celui qui prend plus de 5 points de temps au-delà de sa part de MRR. Les couleurs des clients sont identiques dans les deux camemberts.
