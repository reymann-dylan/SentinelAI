# Sentinel AI pour Microsoft Word

<img align="right" width="350" alt="Capture d'écran de Sentinel AI" src="https://github.com/user-attachments/assets/3e944397-9a49-4791-b745-cb96c6849b2e" />

Sentinel AI est un complément (Add-in) intelligent et ultra-léger pour Microsoft Word. Il ne s'agit pas d'un logiciel lourd, mais d'un outil intégré qui ajoute un assistant IA et un correcteur avancé directement dans votre traitement de texte, avec une priorité absolue donnée à la confidentialité de vos données.

## Qu'est-ce que c'est ?
Fini les allers-retours entre Word et votre navigateur web. Sentinel AI ouvre un panneau latéral qui vous permet de corriger, traduire et reformuler du texte à la volée. 

Vous pouvez également discuter directement avec votre document : l'assistant utilise le texte que vous surlignez comme contexte pour répondre à vos questions avec précision. L'outil est pensé pour s'adapter à votre flux de travail grâce à un système de "Personas" (pour dicter le ton de l'IA, comme "Juridique" ou "Créatif") et des boutons d'action rapide entièrement personnalisables.

## Confidentialité & Sécurité
Sentinel AI a été conçu pour les professionnels et les écrivains qui manipulent des documents sensibles.

Si vous choisissez d'utiliser une IA locale (comme LM Studio ou Ollama), aucune donnée ne part dans le Cloud. L'intégralité de la réflexion de l'IA se fait sur votre propre machine. Une connexion internet est requise uniquement au moment de l'ouverture du panneau pour charger l'interface visuelle. Cependant, absolument aucun mot de votre document ne quitte votre ordinateur : le texte transite strictement en local entre Word et votre IA. 

*(Note : L'outil supporte également les API Cloud classiques comme OpenAI, Mistral ou Groq si vous préférez utiliser la puissance des serveurs en ligne via votre propre clé API).*

## Fonctionnement et stockage des données
L'architecture de ce complément est conçue pour être transparente. L'interface graphique est hébergée de manière sécurisée sur GitHub Pages, ce qui vous garantit d'avoir toujours la dernière version visuelle sans rien télécharger de nouveau. 

Le petit fichier manifeste présent sur votre ordinateur sert uniquement de pont entre Word et cette interface. Concernant vos paramètres (clés API, instructions personnalisées, historique), ils sont sauvegardés exclusivement sur votre disque dur, dans le stockage local du navigateur interne de Microsoft Word. Rien n'est envoyé sur un serveur tiers. Si vous désinstallez l'outil, ces données disparaissent avec lui.

## Guide d'Installation (Release v0.0.1)
L'installation de cet Add-in est très rapide et ne requiert aucune compétence technique.

* Téléchargez le fichier `SentinelAI_v0.0.1.zip` depuis la section Releases de ce projet.
* Extrayez le dossier sur votre ordinateur (par exemple, dans vos Documents).
* Double-cliquez sur le fichier `install.bat`. Si Windows affiche une alerte de protection, cliquez sur "Informations complémentaires" puis sur "Exécuter quand même".
* Ouvrez Microsoft Word : vous trouverez Sentinel AI directement dans l'onglet Compléments.
* Installer LM Studio et instaler via LM Studio léger comme Gemma 4 (model e4b ou e2b) ou encore Mistral 7b
* [! Il est recommender d'avoir une carte graphique moderne sur son ordinateur pour un fonctionnement optimal]

*(Les dossiers contenant le code source sont fournis dans l'archive à titre informatif, mais l'outil utilise la version hébergée en ligne pour fonctionner).*

## Désinstallation
Sentinel AI respecte votre machine. Pour le retirer, fermez Word, retournez dans votre dossier SentinelAI et double-cliquez sur `uninstall.bat`. L'Add-in est alors instantanément déconnecté de Word. Vous pouvez ensuite simplement supprimer le dossier de votre ordinateur.

## Prérequis pour l'IA locale
Pour profiter de l'expérience gratuite et 100% privée, nous recommandons d'utiliser le logiciel LM Studio. Une fois LM Studio installé et votre modèle téléchargé, allez dans l'onglet Local Server, assurez-vous que le port est réglé sur 1234, activez l'option CORS, et démarrez le serveur. Sentinel AI s'y connectera automatiquement.
