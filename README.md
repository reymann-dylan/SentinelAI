# Sentinel AI pour Microsoft Word

<img align="right" width="350" alt="Capture d'écran de Sentinel AI" src="https://github.com/user-attachments/assets/3e944397-9a49-4791-b745-cb96c6849b2e" />

Sentinel AI est un complément (Add-in) "intelligent" et ultra-léger pour Microsoft Word. Il ne s'agit pas d'un logiciel lourd, mais d'un outil intégré qui ajoute un assistant IA et un correcteur avancé directement dans votre traitement de texte, avec une priorité absolue donnée à la confidentialité de vos données.
_*(Note : conçu en un week end, il ne s'agit que d'un proof of concept, les mis a jours seront très éparces, libre a vous de vous "approprier" le projet)_

## Confidentialité & Sécurité
Sentinel AI a été conçu pour manipulent des documents sensibles localement.

Si vous choisissez d'utiliser une IA locale (comme LM Studio ou Ollama), aucune donnée ne part dans le Cloud. L'intégralité de la réflexion de l'IA se fait sur votre propre machine. Une connexion internet est requise uniquement au moment de l'ouverture du panneau pour charger l'interface visuelle qui elle est charger depuis sur git. 

*(Note : L'outil supporte également les API Cloud classiques comme OpenAI, Mistral ou Groq si vous préférez utiliser la puissance des serveurs en ligne via votre propre clé API) (jamais tester personnellement).*

## Fonctionnement et stockage des données
L'architecture de ce complément est conçue pour être transparente. L'interface graphique est hébergée de manière sécurisée sur GitHub Pages, ce qui vous garantit d'avoir toujours la dernière version visuelle sans rien télécharger de nouveau. 

Le petit fichier manifeste présent sur votre ordinateur sert uniquement de pont entre Word et cette interface. Concernant vos paramètres (clés API, instructions personnalisées, historique), ils sont sauvegardés exclusivement sur votre disque dur, dans le stockage local du navigateur interne de Microsoft Word. Rien n'est envoyé sur un serveur tiers. Si vous désinstallez l'outil, ces données disparaissent avec lui.

## Guide d'Installation (Release v0.0.1)
L'installation de cet Add-in est très rapide et ne requiert aucune compétence technique.

* Téléchargez le fichier `SentinelAI_v0.0.1.zip` depuis la section Releases de ce projet.
* Extrayez le dossier sur votre ordinateur (par exemple, dans vos Documents).
* Double-cliquez sur le fichier `install.bat`. Si Windows affiche une alerte de protection, cliquez sur "Informations complémentaires" puis sur "Exécuter quand même".
* Ouvrez Microsoft Word : vous trouverez Sentinel AI directement dans l'onglet Compléments.
* Installez LM Studio et téléchargez un modèle léger comme le Gemma 4 (modèle e4b ou e2b) ou le Mistral 7b.
* [ ! Pour des performances optimales, il est recommandé d'utiliser une carte graphique récente.]

*(Les dossiers contenant le code source sont fournis dans l'archive à titre informatif, mais l'outil utilise la version hébergée en ligne pour fonctionner).*

## Désinstallation
Sentinel AI respecte votre machine. Pour le retirer, fermez Word, retournez dans votre dossier SentinelAI et double-cliquez sur `uninstall.bat`. L'Add-in est alors instantanément déconnecté de Word. Vous pouvez ensuite simplement supprimer le dossier de votre ordinateur.

## Prérequis pour l'IA locale
Pour profiter de l'expérience gratuite et 100% privée, et simple je recommande d'utiliser le logiciel LM Studio. Une fois LM Studio installé et votre modèle téléchargé (penser pour Gemma-4-e2b, e4b & 12b), allez dans l'onglet Local Server, assurez-vous que le port est réglé sur 1234, activez l'option CORS, et démarrez le serveur. Sentinel AI s'y connectera automatiquement.
