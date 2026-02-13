## Ouverture du reporting CA d'une des régions

## VBA : Nettoyage des 5 onglets
- Résumé : Le code nettoie les plages du reporting comprenant des valeurs brutes, en moins de 10 secondes, tout en prenant soin de ne toucher aucune plage contenant des formules.
- Le code annoté est disponible sous le nom "CODE VBA - NETTOYAGE DES DONNEES" en pièce-jointe du projet.

## VBA : Import des nouvelles bases
- Résumé : Le code importe les plages des fichiers "source" vers les plages des fichiers "destination" en 30 sec. Le code est factorisé, scalable et facilement auditable via des logs.
- Le code annoté est disponible sous le nom "CODE VBA - IMPORT DES DONNEES" en pièce-jointe du projet.

## Controle cohérence et actualisation du reporting
- Le contrôle de cohérence s'effectue manuellement après avoir actualisé les données. Il consiste à détecter la présence ou non d'anomalies dans les feuilles de calcul et dans la synthèse envoyée à la direction.
- Automatiser 90% du travail pour garder 10% de contrôles permet de gagner en productivité tout en conservant une maitrise totale du processus.

## VBA : Envoi du reporting
- Le mail est généré avec les destinataires et les pièces-jointes renseignées, de plus l'objet et le corps du mail sont remplis de manière automatique.
- Le brouillon est enregistré dans outlook et pret pour un dernier controle avant envoi.
- Le code annoté est disponible sous le nom "CODE VBA - ENVOI DU MAIL" en pièce-jointe du projet.

## Reporting terminé
