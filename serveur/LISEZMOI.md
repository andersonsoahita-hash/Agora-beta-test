# Serveur Agora test (Google Sheet + Apps Script)

1. Ouvre le Google Sheet, menu **Extensions > Apps Script**.
2. Efface le contenu, colle `Code.gs`, enregistre.
3. Lance la fonction **installer** (bouton Exécuter) et accepte les autorisations (le compte du Sheet doit être celui qui envoie les e-mails).
4. **Déployer > Nouveau déploiement > Application Web** : Exécuter en tant que *Moi*, accès *Tout le monde*. Copie l'URL « /exec ».
5. Recharge le Sheet : un menu **Agora test** apparaît (Valider, Bloquer, Débloquer, Supprimer).
6. Donne l'URL /exec : elle sera branchée dans l'appli de test.

Dans l'onglet **Comptes** : colonne *Expire le* = date de fin d'accès (laisser vide = sans limite).
