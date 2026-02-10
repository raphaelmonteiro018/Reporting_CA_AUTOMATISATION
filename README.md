## Extractions à partir de l'ERP
- Les extractions sont programmées depuis l'ERP et envoyées par mail du lundi au vendredi à heure fixe, exemple :
<img width="1757" height="343" alt="image" src="https://github.com/user-attachments/assets/c18aef5a-ec7f-4ddd-a86d-b025e1e2d07e" />

## Power Automate Desktop : Suppression et remplacement des anciennes extractions
- Flux principal à partir duquel sont appelés les sous-flux :
<img width="1161" height="252" alt="image" src="https://github.com/user-attachments/assets/c78e7015-f283-4435-8cec-513bdc2c47cc" />

- Extrait d'un sous-flux de nettoyage des anciens fichiers / remplacement par les nouveaux :
<img width="1205" height="721" alt="image" src="https://github.com/user-attachments/assets/ccb51dae-c993-4c05-a1b4-56bef037cc49" />

## Power Automate Desktop : Retraitement des extractions (suppression de colonnes et filtrage)
- Extrait d'un sous-flux de traitement :
<img width="1210" height="667" alt="image" src="https://github.com/user-attachments/assets/e6367310-9428-436d-bf36-3c2694424acb" />

## Power Automate Desktop : Résultat final
Le workflow s'exécute en arrière-plan, en 1 minute les extractions des deux régions sont actualisées, rangées et prêtes pour l'import via VBA.

- Voici l'arborescence de dossiers utilisée pour le rangement des extractions sur une région :
<img width="707" height="556" alt="image" src="https://github.com/user-attachments/assets/745197fa-81a7-441a-9ac8-a0b98066805f" />


VBA permet ensuite d'aller piocher les plages désirées dans chaque fichier exporté pour les coller dans les plages spécifiées du reporting.
