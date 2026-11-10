# Agora · version test (by Stephen)

Copie de l'application Agora pour une phase de test. Stockage isolé de l'application principale.
Le serveur (comptes, contenus, suivi) sera relié ensuite.

## Couper ou limiter l'accès (sans serveur)
Modifier `acces.json` sur GitHub (icône crayon) puis enregistrer ; c'est pris en compte en une à deux minutes, à la prochaine ouverture de l'appli avec du réseau :
- `"actif": false` : écran « Accès suspendu » avec ton message.
- `"jusqu_au": "2026-11-24"` : l'accès se ferme tout seul après cette date.
Limite : c'est un interrupteur pour tous, pas par personne, et il bloque l'écran sans effacer les données. Le vrai blocage par compte viendra avec le serveur.
