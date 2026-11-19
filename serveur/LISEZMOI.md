# Serveur Agora test (Google Sheet + Apps Script)

1. Ouvre le Google Sheet, menu **Extensions > Apps Script**.
2. Efface le contenu, colle `Code.gs`, enregistre.
3. Lance la fonction **installer** (bouton Exécuter) et accepte les autorisations (le compte du Sheet doit être celui qui envoie les e-mails).
4. **Déployer > Nouveau déploiement > Application Web** : Exécuter en tant que *Moi*, accès *Tout le monde*. Copie l'URL « /exec ».
5. Recharge le Sheet : un menu **Agora test** apparaît (Valider, Bloquer, Débloquer, Supprimer).
6. Donne l'URL /exec : elle sera branchée dans l'appli de test.



## Notifications (défi du jour, série, relance)
1. Crée un compte gratuit sur onesignal.com, puis une app « Web » (Typical Site). Site : l'adresse GitHub Pages d'Agora test, icône : `icons/icon-512.png`.
2. Copie l'**App ID** (à me donner : il est public) et la **REST API Key** (secrète, ne la partage nulle part sauf ici).
3. Apps Script > Paramètres du projet > Propriétés du script : `OS_APP_ID` et `OS_API_KEY`.
4. Colle le nouveau `Code.gs`, lance `installerNotifs` une fois, puis Déployer > Nouvelle version.
5. Envois automatiques (heure de Paris) : 12 h relance après 3 jours, 18 h défi du jour, 20 h série en jeu, 21 h palier de série (3, 7, 14, 30, 60, 100 jours), et toutes les heures un contrôle de nouvelle version (1 notification par jour au maximum).
