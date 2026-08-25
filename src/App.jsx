import React, { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { ICON_DATA, BILAN_ICON_DATA } from "./iconData";
import { ICON_KEYS, ICON_MAX, BILAN_CARDS } from "./gameData";
import { CARDS_AH } from "./cardsAH";

const PARCOURS = [
  { id: "donnees", label: "Des données sensibles", color: "#f59e0b", cards: [18,25,28,29,35,38,50,61,65,68,76,77,78,81,98] },
  { id: "roles", label: "Des rôles pour l'IA", color: "#10b981", cards: [1,2,3,4,5,6,7,25,29,31,59,62,71,83,84,85,86,93,98] },
  { id: "transparence", label: "Un peu de transparence", color: "#ef4444", cards: [11,17,27,34,45,56] },
  { id: "scribe", label: "L'IA scribe", color: "#8b5cf6", cards: [16,17,20,57,58,89,90,91,95,99,100] },
  { id: "competences", label: "Des compétences à ne pas oublier", color: "#06b6d4", cards: [6,8,14,15,16,20,26,36,46,58,60,66,70,72,80,82,85,87,88,91,96,97,99] },
  { id: "societe", label: "Impact sur la société", color: "#ec4899", cards: [10,24,30,32,33,44,49,55,59,62,66,76,77] },
  { id: "planete", label: "Impact sur la planète", color: "#22c55e", cards: [8,12,13,19,35,51,52,53,54,55,56,67,92] },
  { id: "encadrer", label: "Soutenir et encadrer", color: "#f97316", cards: [1,2,4,6,26,28,31,47,48,57,64,79,80,95] },
  { id: "veracite", label: "Ce résultat est-il juste et vrai ?", color: "#3b82f6", cards: [5,9,14,22,37,39,48,63,64,69,73,74,75,79,100] },
  { id: "prompt", label: "Prompter promptement", color: "#d946ef", cards: [21,22,23,25,70] },
  { id: "esprit_critique", label: "Croire ou vérifier ?", color: "#84cc16", cards: [69,70,71,72,73,74,75,84,94] },
  { id: "vie_etablissement", label: "Au-delà des cours", color: "#0ea5e9", cards: [46,61,82,86,87,88,89,90,92,93,94,96,97] },
]








function getIcons(qid, opt) {
  const keys = ICON_KEYS[qid];
  if (!keys) return [];
  const list = opt === "A" ? keys.a : keys.b;
  return (list || []).map(k => ICON_DATA[k]).filter(Boolean);
}


// ── GRIST CONFIG ────────────────────────────────────────────────────
// Document partagé par les deux apps solo (jeunes/adultes), une table
// par app. Écriture publique (Create uniquement) et lecture publique
// (Read uniquement sur la table de résumé) réglées via les Access
// Rules du document Grist — voir la doc du projet pour le détail.
const GRIST_SERVER = "https://grist.numerique.gouv.fr";
const GRIST_DOC_ID = "7aCYfhin1pd8";
const GRIST_API_BASE = `${GRIST_SERVER}/api/docs/${GRIST_DOC_ID}`;
const GRIST_VOTES_TABLE = "Table1"; // table de votes pour la version jeunes
const GRIST_SUMMARY_TABLE = "COMPLETER_NOM_RESUME_JEUNES"; // voir Raw Data dans Grist
const ADMIN_PASSWORD = "ethiqia2024";

const PROFILES = [
  { value: "cycle4", label: "Cycle 4 (5ème – 3ème)" },
  { value: "lgt", label: "Lycée général et technologique" },
  { value: "lp", label: "Lycée professionnel" },
  { value: "sup", label: "Études supérieures" },
];

const questions = [
  { id: 1, titre: "IA mentor", situation: "J'ai entendu parler d'une IA mentor : elle me sert de coach, en me donnant des commentaires au fur et à mesure de mon travail.", altA: "Oui, c'est plus facile et plus accessible qu'un prof qui a rarement le temps de me guider.", altB: "Non, même si j'ai besoin d'un coach, je préfère tracer mon chemin par mes propres moyens." },
  { id: 2, titre: "IA tuteur", situation: "J'ai entendu parler d'une IA tuteur : elle prendrait en charge mon apprentissage, à la place du prof.", altA: "Pourquoi pas, le prof n'est pas toujours disponible pour m'aider.", altB: "Surtout pas, ça reste une machine, elle ne me connaît pas et elle demande trop de données personnelles." },
  { id: 3, titre: "IA équipière", situation: "J'ai entendu parler d'une IA équipière : dans un travail de groupe, elle fait équipe avec nous en donnant son avis.", altA: "Oui, ce serait pas mal d'avoir un autre avis, même si c'est celui d'une machine.", altB: "Non, c'est déjà difficile de s'entendre entre nous, alors avec une voix en plus…" },
  { id: 4, titre: "IA coach", situation: "J'ai entendu parler d'une IA coach : elle me pose des questions sur ce que j'ai fait, sur ce que je pense avoir réussi et elle me donne des conseils pour progresser.", altA: "Oui, j'ai toujours rêvé d'avoir une voix qui m'aiderait à savoir où j'en suis !", altB: "Non, ça va pas du tout ! Elle va collecter plein de données sur mon travail. C'est trop personnel pour lui laisser tout ça !" },
  { id: 5, titre: "IA simulateur", situation: "J’utilise une IA-simulateur.Elle reproduit une situation, comme si c'était réel. Ça fait jeu vidéo, et c'est motivant.", altA: "Oui, c'est quand même mieux qu'une étude de cas avec plein de questions.", altB: "Non, une simulation reste une simulation : elle ne représente pas assez bien ce qu'on voit dans la vraie vie." },
  { id: 6, titre: "IA outil", situation: "J’utilise une IA-outil. Elle est spécialisée dans une tâche précise (programmation, travail sur des textes, génération d'images…).", altA: "Oui,  elle sera plus efficace sur ce qu'on cherche à faire.", altB: "Non, on a besoin de travailler ces compétences. Avec l'IA, on ne pourra jamais les entraîner correctement." },
  { id: 7, titre: "IA soutien", situation: "Cette IA est là pour aider les élèves qui ont des difficultés bien identifiées. Elle leur sert de \"béquille\" pour faire ce qu'ils ne seraient pas capables de faire.", altA: "Oui, sans IA, il est bien difficile de leur apporter de l'aide.", altB: "Non, ce n'est pas juste pour les autres qui ne bénéficieront pas de cette IA." },
  { id: 8, titre: "Recherche documentaire", situation: "Pour une recherche documentaire générale, j’utilise une IA générative, c'est bien plus rapide !", altA: "Une IA générative, comme cela, on passe plus rapidement au travail de compréhension, de liens, de traitement des données", altB: "Un moteur de recherche, mais ça demande de critiquer les résultats et de  sélectioner les informations utiles, avec leurs sources." },
  { id: 9, titre: "Un résultat préoccupant", situation: "J'ai utilisé l'IA, mais je trouve que le résultat n'est pas à la hauteur de ce que j'espérais.", altA: "Je retente, jusqu'à ce que le résultat soit satisfaisant.", altB: "Je retente mais pas trop, il paraît que ça consomme beaucoup d'énergie et de ressources." },
  { id: 10, titre: "Justice ?", situation: "J'ai fait seul mon devoir à la maison et je l'ai rendu, mais j'ai une moins bonne note que d'autres qui ont utilisé l'IA.", altA: "Je réagis. Il faut que les profs soient capables de gérer cette injustice.", altB: "Je ne fais rien. Ça se reproduira, et ça ne va pas m'encourager à continuer comme ça." },
  { id: 11, titre: "Confiance ?", situation: "J'ai terminé d'écrire un texte. Le prof se montre méfiant. Il pense que j'ai utilisé une IA car c'est meilleur que d'habitude.", altA: "Je continuerai à faire comme ça. Je n'ai rien à me reprocher.", altB: "J'utilise un site qui permet de vérifier si mon texte est ressemblant à ce qu'une IA pourrait produire." },
  { id: 12, titre: "Gaspillage", situation: "J'ai testé l'outil ComparIA. J'ai vu toutes les ressources que j'ai utilisées pour mes demandes à l'IA. C'est énorme…", altA: "Tant pis, si je me préoccupe de cela, je ne ferai plus grand-chose.", altB: "Quand même, ça me gêne. Je vais essayer de trouver des pratiques un peu plus responsables." },
  { id: 13, titre: "Grand ou petit modèle ?", situation: "Je privilégie les grands modèles d’IA (Gemini, ChatGPT, Mistral, Claude…), plutôt que les modèles hébergés en local dans mon établissement.", altA: "Oui, j'ai l'habitude avec ces outils. C'est plus simple.", altB: "Non, il paraît que les modèles en local sont plus petits, moins consommateurs, mais ils sont moins accessibles." },
  { id: 14, titre: "Où se trouve la source ?", situation: "J'ai fait une recherche avec l'IA, mais on m'a demandé les sources. Et l'IA que j'utilise ne me la donne pas.", altA: "Tant pis, on fera sans, ce n'est pas si important.", altB: "Quand même, je vais utiliser un moteur de recherche ou des livres, ça existe encore." },
  { id: 15, titre: "Les deux à la fois ?", situation: "J'ai utilisé un moteur de recherche, et j'ai vu qu'il proposait d'utiliser une IA pour avoir un résumé des résultats.", altA: "C'est parfait, je ne suis plus obligé de cliquer sur les liens et de lire les résultats !", altB: "C'est limité, je ne peux même plus faire mon choix d'informations moi-même !" },
  { id: 16, titre: "Aide à l'écriture", situation: "Je dois écrire un texte. Quel travail vais-je demander à l'IA ?", altA: "Quitte à faire, tout le travail à ma place.", altB: "De la reformulation du texte que j'ai écrit moi-même." },
  { id: 17, titre: "Transparence", situation: "Je dois rendre un texte. Je l'ai fait en partie avec de l'IA.", altA: "Je rends mon travail, sans aucune précision.", altB: "Je copie et colle ma demande à l'IA, je prof saura comment je l'ai utilisée." },
  { id: 18, titre: "Nourrir une IA", situation: "Je voudrais que l'IA retravaille le fichier que j'ai créé (un exposé, un texte, peu importe).", altA: "Aucun problème, même si l'IA va conserver mon fichier et le réutiliser plus tard.", altB: "Pas question de la laisser utiliser ce que j'ai fait. C'est moi le propriétaire." },
  { id: 19, titre: "Calculatrice de luxe", situation: "Pour faire un calcul, je demande directement à l'IA. C'est moins galère que de sortir la calculatrice.", altA: "Oui, ça permet d'aller plus vite.", altB: "Non, inutile d'utiliser un outil aussi puissant et consommateur pour un truc aussi facile." },
  { id: 20, titre: "Qui fera le résumé ?", situation: "Je dois faire un exposé pour résumer tout un sujet.", altA: "Je demande à l'IA de faire la synthèse de ce que j'ai trouvé. Rapide et efficace.", altB: "Je fais la synthèse moi-même, sinon, je ne comprendrai rien au sujet." },
  { id: 21, titre: "Ultra rapide", situation: "Je dois me dépêcher. J'ai une demande à faire à l'IA.", altA: "Je fais la demande comme sur un moteur de recherche, quitte à l'améliorer par la suite.", altB: "Je prends le temps de bien faire ma première demande. Ça m'évitera d'en faire d'autres pour préciser la réponse." },
  { id: 22, titre: "Doutes", situation: "J'ai fait une recherche avec l'IA et j'ai des doutes sur les résultats.", altA: "Je compare avec les résultats d'autres IA.", altB: "Je reste sur ces résultats. D'autres IA ne m'apporteront rien de plus." },
  { id: 23, titre: "En une fois ou en plein de fois", situation: "Je fais une demande à l'IA. Quelle sera la stratégie que je vais adopter ?", altA: "Je le fais en plusieurs fois, en précisant mes demandes.", altB: "Je travaille ma demande pour avoir un résultat acceptable rapidement." },
  { id: 24, titre: "Théorie du remplacement", situation: "J’ai entendu parler de la théorie du remplacement. Il paraît que de nombreux emplois seront détruits car remplacés par l’IA", altA: "Même pas peur. On trouvera d'autres emplois dans le futur, on s'adaptera.", altB: "ça me stresse. Je ne vais pas commencer à me former à un métier qui pourrait disparaître." },
  { id: 25, titre: "Un peu d'entraînement", situation: "Je demande à l'IA de créer des exercices pour que je puisse réviser. Mais pour être efficace…", altA: "Je lui donne uniquement les cours de mes profs.", altB: "Je lui demande de chercher où elle veut, sans la restreindre." },
  { id: 26, titre: "Intuition de départ", situation: "J'ai un exposé à faire. Je ne sais pas comment commencer.", altA: "Je demande à l'IA de me donner le départ. Je continuerai moi-même ensuite.", altB: "Je veux commencer moi-même, sinon, je ne maîtrise pas le sujet. L'IA pourra m'aider plus tard, pour d'autres tâches." },
  { id: 27, titre: "Absence de cadre", situation: "Dans mon établissement scolaire, il n'y a pas de règle sur l'utilisation de l'IA. Du coup, on ne sait pas si on peut l'utiliser ou pas.", altA: "Je préfère éviter au maximum de l'utiliser, ça pourrait toujours m'être reproché.", altB: "Je l'utilise quand même, s'il n'y a pas de cadre, c'est qu'il n'y a pas de problème." },
  { id: 28, titre: "Coaching par l'IA", situation: "J'ai des difficultés et on me propose de tester une IA qui va me coacher.", altA: "Oui, ça peut toujours m'aider.", altB: "Non, il lui faudra plein de données sur moi, et je n'ai pas envie de ça." },
  { id: 29, titre: "L'IA va m'orienter", situation: "Une IA conseillère d'orientation est en expérimentation. Elle serait capable de m'orienter, et de participer aux procédures d'admission.", altA: "Oui, ça ne sera pas pire que les jurys qui ne sont pas toujours justes.", altB: "Non, c'est mon avenir et mes données, je ne veux pas qu'une IA puisse influencer mon destin." },
  { id: 30, titre: "Une IA magique", situation: "Une IA agentique est capable de me donner des réponses en faisant plusieurs tâches l'une après l'autre. C'est pratique : je donne ce qu'il y a à faire, et je récupère le résultat !", altA: "Oui, c'est tellement plus rapide que de décomposer toutes ces tâches !", altB: "Non, je n'ai aucun moyen de vérifier et de valider ce qu'elle fait, et de corriger s'il y a besoin." },
  { id: 31, titre: "Chatbot motivationnel", situation: "On m'a proposé de tester un chatbot motivationnel. Je parle avec lui quand je suis démotivé, et il m'encourage.", altA: "Oui, ça aide quand on ne trouve personne pour se rebooster.", altB: "Non, ça reste une machine programmée, pas un ami ou un psy." },
  { id: 32, titre: "IA sélective 1", situation: "Dans ma classe, on nous propose des comptes pour utiliser une IA pro, mais il n'y a pas assez de comptes pour tout le monde.", altA: "On les réserve pour les personnes les plus bosseuses, pour qu'elles continuent à performer.", altB: "On les réserve aux personnes en difficulté, pour les aider à s'améliorer." },
  { id: 33, titre: "IA sélective 2", situation: "Dans ma classe, on nous propose des comptes pour utiliser une IA pro, mais il n'y a pas assez de comptes pour tout le monde.", altA: "On les réserve pour une partie de la classe, tant pis pour l'égalité.", altB: "Soit c'est l'IA pour tout le monde, soit c'est l'IA pour personne." },
  { id: 34, titre: "Trop de triche !", situation: "Dans ma classe, il y a trop d'élèves qui utilisent l'iA pour les évals. Les profs s'en rendent compte et cherchent des solutions.", altA: "Ils devraient tout interdire et renforcer les moyens pour le contrôler.", altB: "Ils devraient nous encourager à dire quand et comment on utilise l'IA, notamment pour les DM." },
  { id: 35, titre: "Illustrations", situation: "J'ai besoin d'images pour illustrer un exposé.", altA: "Je demande à une IA générative pour les créer rapidement, selon mes instructions.", altB: "Je demande à un moteur de recherche, surtout que je ne sais pas exactement ce que je voudrais." },
  { id: 36, titre: "Exposé", situation: "Je dois faire un exposé. Je ne peux pas le faire seul, il me faut de l'aide.", altA: "Je préfère l'IA, c'est plus rapide et pratique.", altB: "Je préfère le groupe de travail, c'est plus convivial." },
  { id: 37, titre: "Supériorité", situation: "Je fais une  demande à une IA américaine, pour savoir quels sont les meilleurs scientifiques de l'histoire. Elle me sort 10 noms de savants blancs, parlant anglais, tous des hommes.", altA: "C'est comme ça, c'est possible que ce soit ces personnes qui ont marqué l'histoire des sciences.", altB: "Il doit y avoir un biais, puisque c'est une IA américaine." },
  { id: 38, titre: "Mettre les mains dans le cambouis", situation: "On me propose un cours sur le fonctionnement des IA : les algorithmes, les modèles d'entraînement, etc.", altA: "Inutile. Je ne veux pas devenir informaticien !", altB: "Utile, j'ai besoin de savoir comment ces outils marchent pour mieux comprendre leurs résultats." },
  { id: 39, titre: "Bleu, blanc, rouge ?", situation: "J'ai le choix entre deux modèles d'IA, qui ont une seule différence : leur nationalité.", altA: "Je préfère le modèle français, même s'il était un peu plus cher.", altB: "Je préfère le modèle américain ou choinois, peut-être plus performant et plus utilisé dans le monde." },
  { id: 40, titre: "Une IA qui signe à ma place ?", situation: "Le collège organise un concours d'écriture. Je peux soumettre un texte largement écrit avec une IA, sous mon nom, ou un texte plus modeste mais entièrement de moi.", altA: "Je soumets le texte assisté par IA. Le règlement ne l'interdit pas explicitement, et le résultat est bien meilleur.", altB: "Je soumets mon propre texte, plus simple. Un concours d'écriture doit récompenser ce qu'on sait vraiment faire soi-même." },
  { id: 41, titre: "Choisir une option « IA-proof » ?", situation: "Pour choisir mes options au lycée, on me conseille de privilégier des matières jugées résistantes à l'IA plutôt que celles que j'aime vraiment.", altA: "Je choisis selon ce conseil. Autant anticiper un monde du travail transformé par l'IA.", altB: "Je choisis selon mes goûts, sans trop y penser. Personne ne sait vraiment ce qui sera épargné par l'IA dans dix ans." },
  { id: 42, titre: "Une IA pour suivre en cas de handicap ?", situation: "Un élève malentendant de ma classe peut utiliser une IA qui sous-titre en direct les cours, mais elle capte aussi parfois des bribes de conversations privées entre élèves.", altA: "Je trouve que c'est un compromis acceptable. L'accessibilité pour lui compte plus que ce risque limité.", altB: "Ça me gêne. Même pour une bonne cause, je n'aime pas savoir que nos conversations peuvent être captées." },
  { id: 43, titre: "Suivre la tendance ou pas ?", situation: "Tous mes amis utilisent une IA pour s'aider à un jeu-concours de culture générale entre classes, ce qui n'est pas franchement dans l'esprit du jeu.", altA: "Je fais pareil qu'eux. Si tout le monde le fait, refuser me désavantage sans rien changer au fond.", altB: "Je refuse de le faire. Même si ça ne change rien au résultat final, je préfère jouer honnêtement." },
  { id: 44, titre: "Gratuit ou payant ?", situation: "Pour m'aider dans mes devoirs, j'ai le choix entre la version gratuite d'une IA (limitée, parfois moins fiable) et un abonnement payant, plus précis et plus rapide.", altA: "Je reste sur la version gratuite. Ça marche déjà bien pour l'essentiel, et tout le monde peut y accéder de la même façon.", altB: "Je prends l'abonnement payant. Les réponses sont plus fiables, ça vaut le coût si j'en ai les moyens." },
  { id: 45, titre: "Elle explique ou elle répond ?", situation: "J'utilise une IA pour comprendre un exercice de maths. Certaines expliquent leur raisonnement étape par étape, d'autres donnent directement le résultat.", altA: "Je préfère celle qui explique son raisonnement. C'est plus long à lire, mais je comprends vraiment la méthode.", altB: "Je préfère celle qui donne directement la réponse. C'est plus rapide, et je peux avancer plus vite dans mes devoirs." },
  { id: 46, titre: "Le manuel enrichi", situation: "Mon collège propose de remplacer un manuel papier par une version numérique enrichie par l'IA, avec exercices interactifs et corrections automatiques.", altA: "Je choisis le manuel numérique. C'est plus ludique et interactif, ça donne envie de travailler.", altB: "Je garde le manuel papier. C'est peut-être moins amusant, mais je retiens mieux en écrivant et en cherchant moi-même." },
  { id: 47, titre: "Une journée sans IA", situation: "Mon établissement instaure un jour par semaine sans écran ni IA, pour tous les élèves.", altA: "Je trouve ça bien. Ça m'oblige à mobiliser mes propres compétences, même si c'est parfois frustrant.", altB: "Je trouve ça excessif. Autant utiliser les outils qu'on a à disposition tous les jours, pas qu'un jour sur sept." },
  { id: 48, titre: "Un oral de contrôle", situation: "J'ai rendu un devoir à la maison. Le prof, qui pense que je l'ai fait avec une IA sans le dire, me propose un oral de contrôle pour vérifier que je maîtrise le sujet.", altA: "Je passe l'oral tranquillement. Si j'ai vraiment compris ce que j'ai rendu, ça ne me fait pas peur.", altB: "Je trouve ça injuste d'être soupçonné sans preuve, même si en vrai j'ai un peu utilisé l'IA sans le dire." },
  { id: 49, titre: "Un camarade sans accès", situation: "Un camarade de classe n'a pas Internet chez lui et ne peut donc pas utiliser l'IA pour ses devoirs, contrairement à moi.", altA: "Je lui propose de venir travailler chez moi ou de lui montrer comment je m'en sers, pour qu'on soit à égalité.", altB: "Je continue comme d'habitude. Ce n'est pas à moi de compenser les inégalités d'accès à Internet." },
  { id: 50, titre: "L'outil du kiosque ou l'outil perso ?", situation: "Mon collège met à disposition un outil d'IA labellisé, hébergé en France, mais un peu plus lent. J'ai aussi accès chez moi à un outil grand public, plus rapide et plus puissant.", altA: "J'utilise l'outil du collège. Il est peut-être moins performant, mais mes données restent maîtrisées et il a été vérifié.", altB: "Je préfère l'outil que je connais déjà à la maison. Il est plus rapide et je n'ai pas à apprendre un nouvel outil." },
  { id: 51, titre: "Petit modèle sobre ou grand modèle puissant ?", situation: "Pour un exercice simple, je peux utiliser un petit modèle d'IA, plus sobre en énergie, ou un grand modèle, plus lourd mais parfois plus précis.", altA: "Je choisis le petit modèle. Pour une tâche simple, ça suffit largement et ça consomme moins.", altB: "Je prends le grand modèle par sécurité. Je préfère être sûr d'avoir la meilleure réponse possible." },
  { id: 52, titre: "Un téléphone plus récent pour l'IA embarquée ?", situation: "Mon téléphone actuel ne peut pas faire tourner les nouvelles fonctions d'IA intégrées. Mes parents me proposent d'en racheter un plus récent.", altA: "Je garde mon téléphone actuel. Je peux toujours utiliser l'IA via une appli ou un site, pas besoin de racheter.", altB: "Je change de téléphone. Autant profiter des nouvelles fonctions si on en a l'occasion." },
  { id: 53, titre: "Vidéo générée ou texte pour mon exposé ?", situation: "Pour illustrer mon exposé, je peux demander à une IA de générer une courte vidéo, plus impressionnante mais très gourmande en ressources, ou me contenter d'un texte et d'images fixes.", altA: "Je reste sur le texte et les images. C'est moins spectaculaire, mais ça suffit pour être compris.", altB: "Je génère la vidéo. Ça rend l'exposé plus vivant, et ce n'est qu'une fois." },
  { id: 54, titre: "Un accès illimité, alors j'en profite ?", situation: "Depuis que mon abonnement IA est illimité et gratuit, je m'en sers pour absolument tout, même des questions que je pourrais résoudre seul en trente secondes.", altA: "Je me limite quand même aux vraies questions utiles. Ce n'est pas parce que c'est gratuit qu'il faut en abuser.", altB: "Je m'en sers pour tout, tant que c'est illimité. Autant profiter de ce que ça permet." },
  { id: 55, titre: "Un centre de données près de chez moi", situation: "Un centre de données va être construit près de ma ville. Il devrait créer des emplois, mais il va aussi consommer beaucoup d'eau et d'électricité localement.", altA: "Je soutiens le projet. Ça crée de l'activité économique, et il faut bien héberger ces services quelque part.", altB: "Je m'inquiète pour les ressources locales : l'eau et l'électricité risquent de manquer pour les habitants." },
  { id: 56, titre: "Une IA muette sur sa consommation", situation: "L'IA que j'utilise ne dit jamais combien elle consomme d'énergie ou d'eau pour répondre à mes demandes. Une autre, moins connue, affiche une estimation à chaque réponse.", altA: "Je choisis celle qui affiche sa consommation, même si elle est moins pratique à utiliser au quotidien.", altB: "Je reste sur celle que je connais. Ce chiffre ne changerait de toute façon pas grand-chose à mes usages." },
  { id: 57, titre: "Une IA qui résume nos décisions", situation: "Au conseil de vie collégienne, on propose d'utiliser une IA pour résumer les débats et proposer une synthèse des décisions à prendre.", altA: "Je trouve ça utile. Ça permet de ne rien oublier et de gagner du temps sur la rédaction.", altB: "Je préfère qu'on rédige nous-mêmes. Une synthèse automatique risque de déformer ce qu'on a vraiment dit." },
  { id: 58, titre: "Une IA prend des notes à ma place", situation: "Pendant les réunions d'un club dont je suis responsable, je peux laisser une IA transcrire et résumer automatiquement, plutôt que de prendre des notes moi-même.", altA: "Je laisse l'IA transcrire. Je peux me concentrer entièrement sur la discussion pendant la réunion.", altB: "Je prends mes notes moi-même. Ça m'oblige à vraiment être attentif à ce qui compte." },
  { id: 59, titre: "Une décision qui me concerne, avec l'aide d'une IA", situation: "Pour m'aider à choisir mon orientation, le CIO utilise un outil d'IA qui analyse mon profil et propose des pistes, toujours validées ensuite par un conseiller humain.", altA: "Je trouve ça rassurant. L'IA élargit les pistes possibles, et un humain valide quand même à la fin.", altB: "Ça me met mal à l'aise. Je préfère que ce soit une personne qui réfléchisse à mon profil depuis le début." },
  { id: 60, titre: "Se former vite, quitte à moins bien ?", situation: "On me propose une formation accélérée à l'IA en une demi-journée, plutôt qu'un module plus complet étalé sur plusieurs semaines.", altA: "Je prends la formation courte. Ça me donne au moins une base, et je pourrai compléter plus tard si besoin.", altB: "Je préfère attendre le module complet. Une demi-journée ne suffit pas pour vraiment comprendre les enjeux." },
  { id: 61, titre: "Construire notre propre outil ou en acheter un ?", situation: "Mon collège hésite entre développer son propre outil d'IA avec les élèves volontaires, ou acheter un abonnement à une solution déjà toute prête.", altA: "Je préfère qu'on achète une solution toute prête. C'est plus fiable et plus rapide à mettre en place.", altB: "Je préfère qu'on construise notre outil. On comprend mieux comment ça marche, même si c'est plus long." },
  { id: 62, titre: "Un ami virtuel pour mes soucis", situation: "Quand j'ai un souci personnel, je peux en parler tout de suite à une IA disponible immédiatement, ou attendre de pouvoir en parler à quelqu'un de mon entourage.", altA: "Je préfère en parler tout de suite à l'IA. Au moins je ne reste pas seul avec mes questions en attendant.", altB: "J'attends d'en parler à quelqu'un que je connais. Une vraie personne comprend mieux ce que je vis." },
  { id: 63, titre: "Signalé à tort par un détecteur d'IA", situation: "Un logiciel censé détecter les textes écrits par IA signale mon devoir comme suspect, alors que je l'ai entièrement rédigé moi-même.", altA: "Je conteste calmement auprès du prof, en expliquant ma méthode de travail, même si ça prend du temps.", altB: "Je laisse couler. Protester risque de me faire passer pour quelqu'un qui a quelque chose à cacher." },
  { id: 64, titre: "Un devoir surveillé par IA", situation: "Pour un devoir à la maison noté, l'établissement propose d'utiliser un logiciel qui surveille mon écran pendant que je compose, pour vérifier que je n'utilise pas d'IA.", altA: "J'accepte la surveillance. Ça garantit que la note reflète vraiment mon propre travail.", altB: "Ça me dérange d'être observé comme ça. Je préfère qu'on me fasse confiance ou qu'on évalue autrement." },
  { id: 65, titre: "Partager mes travaux pour entraîner une IA", situation: "Une IA éducative gratuite propose d'être encore meilleure si les élèves acceptent que leurs travaux et corrections servent à l'entraîner.", altA: "J'accepte. Si ça permet d'améliorer l'outil pour tout le monde, ça me semble un juste échange.", altB: "Je refuse. Mon travail m'appartient, je ne veux pas qu'il serve à entraîner un produit commercial." },
  { id: 66, titre: "Traduire ou apprendre la langue ?", situation: "Un nouvel élève ne parle pas encore bien français. Je peux utiliser une IA pour traduire nos échanges en direct, ou essayer de communiquer avec les quelques mots qu'il connaît déjà.", altA: "J'utilise l'IA pour traduire. On peut vraiment se comprendre tout de suite, sans que ça prenne des mois.", altB: "J'essaie de communiquer autrement, avec des gestes et le peu de mots qu'il connaît. C'est plus lent, mais on progresse ensemble." },
  { id: 67, titre: "Se chauffer grâce aux serveurs IA", situation: "Le centre de données voisin propose de récupérer la chaleur produite par ses serveurs pour chauffer une partie de mon lycée, gratuitement, en échange d'un partenariat de long terme avec l'exploitant.", altA: "Je trouve que c'est une bonne idée. Autant récupérer une chaleur qui serait perdue de toute façon.", altB: "Ça me gêne qu'on devienne dépendant d'un partenariat avec une entreprise dont l'activité pollue par ailleurs." },
  { id: 68, titre: "IA sur mon appareil ou dans le cloud ?", situation: "Pour utiliser l'IA sans connexion et sans envoyer mes données à l'extérieur, mon téléphone propose une version installée directement dessus, moins puissante qu'une version en ligne.", altA: "Je choisis la version installée sur mon téléphone. Mes données ne sortent pas de l'appareil, c'est plus rassurant.", altB: "Je choisis la version en ligne. Elle est plus puissante, et je fais confiance au service que j'utilise." },
  { id: 69, titre: "Elle cite une source, mais dit-elle vrai ?", situation: "Une IA me donne une réponse avec une source qui a l'air sérieuse. En vérifiant, je trouve bien la source, mais elle ne dit pas exactement ce que l'IA prétend.", altA: "Je fais confiance à la réponse de l'IA. La source existe, ça me suffit comme preuve.", altB: "Je me méfie et je signale l'écart. Une source qui existe ne veut pas dire que l'IA l'a bien comprise." },
  { id: 70, titre: "Chercher les mots ou chercher la réponse ?", situation: "Je bloque sur un sujet de recherche parce que je ne trouve pas les bons mots-clés à taper dans un moteur de recherche. Une IA peut soit me donner directement la réponse toute faite, soit juste m'aider à trouver des mots-clés à utiliser ensuite moi-même.", altA: "Je lui demande directement la réponse. C'est plus simple et plus rapide.", altB: "Je lui demande juste des mots-clés, puis je fais ma propre recherche avec de vraies sources." },
  { id: 71, titre: "Elle dit bonjour, je réponds ?", situation: "Certains de mes camarades disent bonjour et merci à l'IA, comme s'ils parlaient à une vraie personne. D'autres trouvent que ça n'a aucun sens puisque ce n'est qu'un programme.", altA: "Je fais pareil, par politesse. Ça ne coûte rien et ça garde de bonnes habitudes.", altB: "Je ne le fais pas. Pour moi, ça entretient l'idée que l'IA comprend ou ressent quelque chose, ce qui n'est pas le cas." },
  { id: 72, titre: "J'ai la réponse, ai-je vraiment compris ?", situation: "Après avoir utilisé une IA pour un exercice difficile, j'ai l'impression d'avoir tout compris. Mais je me demande si je saurais refaire l'exercice tout seul, sans elle.", altA: "Je passe à la suite. Si la réponse est juste et que je comprends l'explication en la lisant, c'est suffisant.", altB: "Je referme l'IA et j'essaie de refaire l'exercice seul, pour vérifier que j'ai vraiment compris et pas juste lu." },
  { id: 73, titre: "Je vérifie, ou je fais confiance ?", situation: "La plupart du temps, je me dis que je devrais vérifier les réponses de l'IA, mais en pratique, je les accepte presque toujours telles quelles.", altA: "Je continue comme ça. Vérifier systématiquement prendrait un temps que je n'ai pas à chaque fois.", altB: "Je décide de vraiment vérifier, au moins une fois sur deux, même si c'est plus long." },
  { id: 74, titre: "Elle hallucine, mais elle a l'air sûre d'elle", situation: "Une IA me donne une réponse fausse, mais formulée avec la même assurance que si elle était vraie. Rien dans le ton ne permet de deviner qu'elle se trompe.", altA: "Je fais quand même confiance au ton assuré. Si elle se trompait, ça se sentirait plus.", altB: "Je me méfie du ton assuré justement. Une IA parle pareil, qu'elle ait raison ou tort." },
  { id: 75, titre: "Une synthèse ou des sources qui se contredisent ?", situation: "Pour mon exposé, je peux demander à l'IA une synthèse toute faite du sujet, ou aller chercher plusieurs sources qui ne disent pas exactement la même chose, et me faire mon propre avis.", altA: "Je prends la synthèse de l'IA. C'est cohérent, bien écrit, et ça m'évite de me perdre dans des sources contradictoires.", altB: "Je vais chercher plusieurs sources, même contradictoires. Ça me permet de comprendre qu'il y a débat, pas une seule vérité." },
  { id: 76, titre: "Un tableau de bord sur mon risque de décrochage", situation: "Le collège utilise un outil d'IA qui repère les élèves à risque de décrochage à partir de mes absences, mes notes et mes résultats, pour alerter plus tôt les adultes qui peuvent m'aider.", altA: "Je trouve ça rassurant. Si je traverse une mauvaise période, quelqu'un peut le remarquer avant que ça s'aggrave.", altB: "Ça me met mal à l'aise. Je n'ai jamais donné mon accord pour que mes données soient analysées comme ça." },
  { id: 77, titre: "Partager mes fiches de révision dans un pot commun", situation: "Une plateforme propose de mettre en commun les fiches de révision générées ou corrigées avec l'IA par tous les élèves d'un même niveau, accessibles à tous gratuitement.", altA: "Je partage mes fiches. Si ça aide d'autres élèves, tant mieux, et je profite aussi de celles des autres.", altB: "Je garde mes fiches pour moi. J'ai mis du temps à les préparer, ce n'est pas à moi de les offrir à tout le monde." },
  { id: 78, titre: "Une IA qui a beaucoup appris de moi", situation: "J'utilise depuis un an la même IA pour m'aider dans mes devoirs. Elle connaît maintenant très bien ma façon de travailler et anticipe mes questions.", altA: "Je trouve ça pratique. Elle me comprend de mieux en mieux, comme un vrai assistant personnel.", altB: "Ça m'inquiète un peu. Je ne sais pas exactement ce qu'elle garde en mémoire sur moi ni pour combien de temps." },
  { id: 79, titre: "Un devoir de groupe avec ou sans IA imposée", situation: "Pour un travail de groupe, le prof impose que personne n'utilise l'IA, mais un camarade du groupe le fait quand même en cachette, et le rendu final en profite.", altA: "Je le signale au prof. La règle vaut pour tout le groupe, et il n'est pas juste que le nôtre bénéficie d'un avantage caché.", altB: "Je ne dis rien. Ce n'est pas à moi de rapporter un camarade, et le résultat du groupe est meilleur comme ça." },
  { id: 80, titre: "Une pause numérique collective", situation: "Ma classe propose de tenter, ensemble, une semaine complète sans aucun usage d'IA, y compris pour les devoirs à la maison, pour voir ce que ça change.", altA: "Je participe avec enthousiasme. Ça peut être intéressant de comparer, même si c'est contraignant.", altB: "Je participe, mais à contrecœur. Ça me semble artificiel de s'interdire un outil qu'on utilisera de toute façon après." },
  { id: 81, titre: "Une appli qui suit mes progrès sur le long terme", situation: "Une appli d'apprentissage garde en mémoire mes progrès sur plusieurs années et ajuste les exercices en fonction de tout mon historique scolaire, y compris mes anciens échecs.", altA: "Je trouve ça utile. Elle voit mieux que quiconque où j'ai vraiment progressé et où je bute encore.", altB: "Ça me gêne qu'elle garde en mémoire mes anciens échecs indéfiniment. J'aimerais pouvoir tourner la page." },
  { id: 82, titre: "Une compétition entre IA et élèves", situation: "Le CDI organise un défi : résoudre une énigme documentaire plus vite qu'une IA, en n'utilisant que des sources classiques (livres, bases de données).", altA: "Je participe avec plaisir. C'est un bon moyen de mesurer ce que je sais vraiment faire sans IA.", altB: "Je trouve l'exercice un peu dépassé. Dans la vraie vie, j'aurai toujours accès à l'IA, autant s'entraîner avec." },
  { id: 83, titre: "Des devoirs adaptés à mon rythme, tout le temps", situation: "Une IA adaptative me propose des exercices toujours calibrés exactement à mon niveau du moment, sans jamais me confronter à quelque chose de trop difficile.", altA: "Je trouve ça motivant. Je ne me décourage jamais parce que c'est toujours à ma portée.", altB: "Je préfère parfois affronter des exercices trop difficiles. C'est aussi comme ça qu'on progresse vraiment." },
  { id: 84, titre: "Une IA embarquée dans mes lunettes ou mon casque", situation: "Une nouvelle génération de lunettes connectées peut m'afficher en direct des informations générées par IA pendant que je me promène ou que je visite un musée.", altA: "Ça m'intéresse. Avoir des explications en direct enrichit ce que je vois, sans rien chercher activement.", altB: "Je préfère sans. Je préfère observer et me poser mes propres questions avant de chercher les réponses." },
  { id: 85, titre: "Recommander l'IA à un plus jeune", situation: "Mon petit frère en primaire me demande comment utiliser l'IA que j'utilise pour mes devoirs de collège.", altA: "Je lui montre, en adaptant à son niveau. Autant qu'il apprenne tôt à s'en servir intelligemment.", altB: "Je lui dis d'attendre. À son âge, je pense qu'il vaut mieux d'abord apprendre à faire sans." },
  { id: 86, titre: "Une IA analyse mes matchs", situation: "Mon club de foot propose une appli qui analyse mes matchs filmés avec l'IA pour repérer mes points faibles techniques, sans passer par l'œil de l'entraîneur.", altA: "Je l'utilise avec plaisir. Ça me donne des retours détaillés que l'entraîneur n'a pas le temps de faire pour chacun.", altB: "Je m'en méfie un peu. Je préfère les retours d'un humain qui connaît mon jeu depuis longtemps, même moins précis." },
  { id: 87, titre: "Une IA compose mon accompagnement", situation: "Pour mon morceau de fin d'année en cours de musique, une IA peut composer un accompagnement instrumental complet à partir de ma mélodie.", altA: "Je l'utilise. Le résultat sonne bien plus riche que ce que je saurais composer seul.", altB: "Je compose moi-même mon accompagnement, même s'il sera plus simple. Je veux que le morceau soit entièrement à moi." },
  { id: 88, titre: "Un projet d'arts plastiques généré", situation: "Pour un projet noté en arts plastiques, je peux utiliser une IA générative pour créer les images de base que je retouche ensuite, ou tout dessiner moi-même depuis le début.", altA: "Je pars d'une image générée par IA. Ça me permet d'aller plus loin dans les finitions plutôt que de bloquer sur l'esquisse.", altB: "Je dessine tout moi-même. Même moins abouti, c'est le seul moyen de vraiment progresser en dessin." },
  { id: 89, titre: "Le journal du collège écrit par IA", situation: "Pour le journal du collège, on peut gagner du temps en demandant à une IA de rédiger un premier jet des articles à partir de nos notes d'interview, sans le préciser aux lecteurs.", altA: "On le fait sans le préciser. Le contenu reste basé sur nos vraies interviews, la rédaction n'a pas besoin d'être signalée.", altB: "On précise l'usage de l'IA dans l'ours du journal. Les lecteurs ont le droit de savoir comment les articles sont produits." },
  { id: 90, titre: "Un voyage scolaire planifié par IA", situation: "Pour préparer notre voyage scolaire, les délégués de classe peuvent utiliser une IA pour construire tout l'itinéraire et les activités, ou le préparer eux-mêmes avec le CDI.", altA: "On utilise l'IA. Elle propose un itinéraire complet en quelques minutes, qu'on n'a plus qu'à valider.", altB: "On le prépare nous-mêmes avec le CDI. C'est plus long, mais on apprend à organiser un vrai projet." },
  { id: 91, titre: "Une lettre de motivation pour mon stage", situation: "Pour mon stage de troisième, je peux demander à une IA de rédiger entièrement ma lettre de motivation, ou l'écrire moi-même en m'inspirant juste d'un exemple.", altA: "Je la laisse rédiger l'essentiel. Le résultat est plus soigné que ce que j'écrirais seul.", altB: "Je l'écris moi-même. Une lettre de motivation doit vraiment venir de moi, même si elle est moins bien tournée." },
  { id: 92, titre: "Une IA anti-gaspillage à la cantine", situation: "La cantine teste un outil d'IA qui prédit combien de plats préparer chaque jour à partir des habitudes de réservation, pour réduire le gaspillage alimentaire.", altA: "Je trouve ça positif. Moins de gaspillage, c'est bon pour tout le monde et pour l'environnement.", altB: "Ça m'inquiète un peu. Et si l'IA se trompe un jour et qu'il n'y a plus assez à manger pour tous ?" },
  { id: 93, titre: "Un arbitre IA pour le tournoi interne", situation: "Pour le tournoi sportif inter-classes, le foyer propose d'utiliser une appli d'IA qui arbitre certains matchs à partir de la vidéo, plutôt que des élèves volontaires.", altA: "Je préfère l'IA. Elle ne favorise personne et évite les disputes sur les décisions arbitrales.", altB: "Je préfère des élèves arbitres. Ça reste notre tournoi, entre nous, avec ses imperfections humaines." },
  { id: 94, titre: "Le CDI qui recommande mes lectures", situation: "Le CDI propose un outil qui recommande des livres personnalisés à partir de mes précédents emprunts, plutôt que de demander conseil au professeur documentaliste.", altA: "J'utilise les recommandations de l'outil. Elles correspondent bien à ce que j'aime déjà lire.", altB: "Je préfère demander au professeur documentaliste. Il peut me faire découvrir des choses auxquelles je n'aurais pas pensé." },
  { id: 95, titre: "Le compte-rendu du conseil de discipline", situation: "En tant que délégué, je participe à la rédaction du compte-rendu d'un conseil de discipline. Une IA peut le rédiger à partir de mes notes prises en séance.", altA: "Je la laisse rédiger le compte-rendu. Ça évite les oublis et c'est plus rapide à produire.", altB: "Je le rédige moi-même, phrase par phrase. C'est un document sensible, je veux garder le contrôle total sur ce qu'il dit." },
  { id: 96, titre: "Identifier les espèces en sortie nature", situation: "En sortie nature, je peux utiliser une appli d'IA qui identifie instantanément les plantes et insectes que je photographie, ou essayer d'abord avec le guide botanique papier.", altA: "J'utilise l'appli. C'est immédiat, et je peux identifier beaucoup plus d'espèces en une sortie.", altB: "J'essaie d'abord avec le guide papier. Chercher moi-même m'aide à mieux retenir ce que j'apprends." },
  { id: 97, titre: "Des dialogues générés pour la pièce de théâtre", situation: "Pour la pièce du club théâtre, on peut demander à une IA de générer des dialogues à partir de notre synopsis, ou les écrire nous-mêmes en atelier d'écriture collective.", altA: "On utilise l'IA pour un premier jet, qu'on retravaille ensuite ensemble. Ça débloque plus vite l'écriture.", altB: "On écrit tout nous-mêmes en atelier. C'est plus lent, mais le texte final nous ressemble vraiment." },
  { id: 98, titre: "Un journal intime commenté par une IA", situation: "Une appli de journal intime propose de commenter mes entrées avec une IA bienveillante, qui répond à ce que j'écris comme le ferait un confident.", altA: "Je trouve ça réconfortant d'avoir un retour, même s'il vient d'une IA, plutôt que d'écrire dans le vide.", altB: "Je préfère qu'il reste vraiment intime, sans aucun retour, même bienveillant. Ça change tout de savoir qu'on me répond." },
  { id: 99, titre: "Une IA qui démarche les entreprises pour mon stage", situation: "Pour trouver mon stage de troisième, une IA peut identifier et contacter automatiquement des entreprises proches de chez moi en mon nom, ou je peux le faire moi-même, appel par appel.", altA: "Je la laisse faire. Elle contacte beaucoup plus d'entreprises que je ne pourrais le faire seul.", altB: "Je le fais moi-même. C'est stressant, mais c'est aussi ma première vraie expérience de démarche professionnelle." },
  { id: 100, titre: "Mon auto-évaluation de fin d'année", situation: "Pour le bilan de fin d'année à remplir avec mon professeur principal, une IA peut rédiger une auto-évaluation soignée à partir de mes notes et de quelques mots-clés que je lui donne.", altA: "Je la laisse rédiger l'essentiel. Le résultat est plus clair que ce que j'écrirais moi-même.", altB: "Je l'écris moi-même, même maladroitement. Ce bilan doit refléter comment je vois vraiment mon année." },
];

const emptyCount = () => ({ A: 0, B: 0 });

async function fetchCounts() {
  // Lit la table de résumé (pré-agrégée par Grist, en lecture publique).
  // On ne lit jamais les votes bruts individuels depuis le client.
  const res = await fetch(`${GRIST_API_BASE}/tables/${GRIST_SUMMARY_TABLE}/records`);
  if (!res.ok) throw new Error("Erreur de chargement");
  const { records } = await res.json();
  const counts = {};
  for (const rec of records) {
    const { NumCarte, Reponse, count } = rec.fields;
    if (!counts[NumCarte]) counts[NumCarte] = emptyCount();
    counts[NumCarte][Reponse] = count;
  }
  return counts;
}

async function submitVote(question_id, option) {
  // Chaque vote est une ligne indépendante (jamais de compteur partagé
  // à lire-puis-réécrire) : élimine par construction le risque de
  // concurrence qui existait avec l'ancien upsertCount côté Supabase.
  const res = await fetch(`${GRIST_API_BASE}/tables/${GRIST_VOTES_TABLE}/records`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ records: [{ fields: { NumCarte: question_id, Reponse: option } }] }),
  });
  if (!res.ok) throw new Error("Erreur de sauvegarde");
}

// PDF generation using jsPDF (loaded from CDN)
function generatePDF(counts, voted, profile) {
  return new Promise((resolve) => {
    if (!window.jspdf) { alert("PDF non disponible, réessayez dans un instant."); return; }
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    const W = 210; const margin = 15; const colW = W - margin * 2;
    const colorDark = [15, 12, 41]; const colorA = [79, 195, 247]; const colorB = [206, 147, 216];
    const colorGray = [120, 110, 100]; const colorLight = [240, 235, 224];

    let y = 0;
    const checkPage = (needed = 20) => { if (y + needed > 275) { doc.addPage(); y = 15; } };

    // Header
    doc.setFillColor(...colorDark);
    doc.rect(0, 0, W, 28, "F");
    doc.setFontSize(22); doc.setFont("helvetica", "bold");
    doc.setTextColor(255, 215, 100);
    doc.text("Ethiq·IA", margin, 14);
    doc.setFontSize(10); doc.setFont("helvetica", "normal");
    doc.setTextColor(...colorLight);
    doc.text("Résultats des votes", margin, 21);
    if (profile) {
      const pl = PROFILES.find(p => p.value === profile);
      doc.text(`Profil : ${pl ? pl.label : profile}`, W - margin, 21, { align: "right" });
    }
    doc.text(`Exporté le ${new Date().toLocaleDateString("fr-FR")}`, W - margin, 14, { align: "right" });

    y = 35;

    // Stats summary box
    const totalA = Object.values(counts).reduce((s, c) => s + (c.A || 0), 0);
    const totalB = Object.values(counts).reduce((s, c) => s + (c.B || 0), 0);
    const totalV = totalA + totalB;
    doc.setFillColor(30, 25, 60);
    doc.roundedRect(margin, y, colW, 18, 3, 3, "F");
    doc.setFontSize(9); doc.setFont("helvetica", "bold");
    doc.setTextColor(...colorLight);
    doc.text(`Total votes : ${totalV}`, margin + 8, y + 7);
    doc.setTextColor(...colorA);
    doc.text(`Option 1 : ${totalA}`, margin + 50, y + 7);
    doc.setTextColor(...colorB);
    doc.text(`Option 2 : ${totalB}`, margin + 90, y + 7);
    doc.setTextColor(...colorGray);
    doc.setFont("helvetica", "normal");
    doc.text(`Questions répondues : ${Object.keys(voted).length} / ${questions.length}`, margin + 130, y + 7);
    y += 24;

    // Cards
    questions.forEach((q) => {
      const c = counts[q.id] || emptyCount();
      const tot = c.A + c.B;
      const uv = voted[q.id];
      const pctA = tot > 0 ? Math.round((c.A / tot) * 100) : null;

      // estimate height needed
      const situLines = doc.splitTextToSize(q.situation, colW - 14);
      const optALines = doc.splitTextToSize(`Option 1 : ${q.altA}`, colW - 14);
      const optBLines = doc.splitTextToSize(`Option 2 : ${q.altB}`, colW - 14);
      const blockH = 10 + (situLines.length * 4.5) + (optALines.length * 4.5) + (optBLines.length * 4.5) + (tot > 0 ? 8 : 0) + 8;
      checkPage(blockH);

      // card bg
      doc.setFillColor(25, 22, 50);
      doc.roundedRect(margin, y, colW, blockH, 3, 3, "F");

      // number badge
      const badgeColor = uv === "A" ? colorA : uv === "B" ? colorB : [80, 70, 120];
      doc.setFillColor(...badgeColor);
      doc.circle(margin + 7, y + 6, 4.5, "F");
      doc.setFontSize(7); doc.setFont("helvetica", "bold");
      doc.setTextColor(255, 255, 255);
      doc.text(String(q.id), margin + 7, y + 7.5, { align: "center" });

      // titre
      doc.setFontSize(10); doc.setFont("helvetica", "bold");
      doc.setTextColor(...colorLight);
      doc.text(q.titre, margin + 14, y + 7);

      // voted badge
      if (uv) {
        doc.setFillColor(...(uv === "A" ? colorA : colorB));
        doc.roundedRect(W - margin - 28, y + 2.5, 26, 7, 2, 2, "F");
        doc.setFontSize(7); doc.setFont("helvetica", "bold");
        doc.setTextColor(255, 255, 255);
        doc.text(`Votre choix : Opt. ${uv === "A" ? "1" : "2"}`, W - margin - 15, y + 7, { align: "center" });
      }

      let cy = y + 12;

      // situation
      doc.setFontSize(8); doc.setFont("helvetica", "italic");
      doc.setTextColor(...colorGray);
      situLines.forEach(l => { doc.text(l, margin + 7, cy); cy += 4.5; });
      cy += 1;

      // opt A
      doc.setFontSize(8); doc.setFont("helvetica", "normal");
      doc.setTextColor(uv === "A" ? colorA[0] : colorLight[0], uv === "A" ? colorA[1] : colorLight[1], uv === "A" ? colorA[2] : colorLight[2]);
      optALines.forEach((l, i) => {
        doc.text((i === 0 ? "Option 1 : " : "           ") + (i === 0 ? q.altA.substring(0, l.length - 9) : l), margin + 7, cy);
        cy += 4.5;
      });

      // opt B
      doc.setTextColor(uv === "B" ? colorB[0] : colorLight[0], uv === "B" ? colorB[1] : colorLight[1], uv === "B" ? colorB[2] : colorLight[2]);
      optBLines.forEach((l, i) => {
        doc.text((i === 0 ? "Option 2 : " : "           ") + (i === 0 ? q.altB.substring(0, l.length - 9) : l), margin + 7, cy);
        cy += 4.5;
      });

      // vote bar
      if (tot > 0 && pctA !== null) {
        const barW = colW - 14; const barH = 2.5; const bx = margin + 7;
        doc.setFillColor(50, 45, 80);
        doc.roundedRect(bx, cy, barW, barH, 1, 1, "F");
        if (pctA > 0) { doc.setFillColor(...colorA); doc.roundedRect(bx, cy, barW * pctA / 100, barH, 1, 1, "F"); }
        doc.setFontSize(7); doc.setTextColor(...colorGray);
        doc.text(`Opt.1: ${c.A} (${pctA}%)`, bx, cy + barH + 4);
        doc.text(`Opt.2: ${c.B} (${100 - pctA}%)`, bx + barW, cy + barH + 4, { align: "right" });
        cy += 8;
      }

      y = y + blockH + 4;
    });

    // Page numbers
    const pageCount = doc.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
      doc.setPage(i);
      doc.setFontSize(8); doc.setFont("helvetica", "normal");
      doc.setTextColor(...colorGray);
      doc.text(`Page ${i} / ${pageCount}  —  Éthiq·IA`, W / 2, 290, { align: "center" });
    }

    doc.save("ethiqia_résultats.pdf");
    resolve();
  });
}


// ── TUTORIAL COMPONENT ──────────────────────────────────────────────
const TUTORIAL_STEPS = [
  {
    emoji: "👋",
    title: "Bienvenue dans Éthiq·IA !",
    color: "#ffd764",
    text: "Un jeu sérieux pour explorer les questions d'éthique autour de l'intelligence artificielle. 6 règles simples à connaître avant de commencer.",
    visual: "welcome",
  },
  {
    emoji: "⚖️",
    title: "L'éthique, ce n'est pas la morale",
    color: "#ce93d8",
    quote: "L'éthique n'est pas la morale : il n'y a pas forcément de meilleur choix. C'est votre propre appréciation de la situation que nous souhaitons.",
    text: "Ne répondez pas en réfléchissant à la réponse attendue ! Votre ressenti compte plus que la bonne réponse.",
    visual: "ethics",
  },
  {
    emoji: "🃏",
    title: "Face à des situations concrètes",
    color: "#4fc3f7",
    text: "Vous ferez face à 39 situations d'utilisation de l'IA. Pour chacune, deux réponses possibles seulement — vous devrez en choisir une seule.",
    sub: "Pas de bonne ou mauvaise réponse. Juste votre point de vue du moment.",
    visual: "cards",
  },
  {
    emoji: "🗺️",
    title: "Deux façons de voter",
    color: "#69f0ae",
    text: "Vous pouvez entrer vos choix de deux façons — c'est la même chose, choisissez celle qui vous convient !",
    items: [
      { icon: "🃏", tab: "Cartes", desc: "Une grille de 39 cartes numérotées — cliquez sur un numéro pour ouvrir la question." },
      { icon: "📋", tab: "Liste", desc: "Toutes les questions déroulées, avec la situation et les deux options visibles." },
    ],
    visual: "nav",
  },
  {
    emoji: "🏅",
    title: "Gagnez des icônes mystères",
    color: "#ffd764",
    text: "À chaque vote, vous gagnez de curieuses icônes sans savoir ce qu'elles signifient. À la fin, consultez l'onglet Bilan pour :",
    items: [
      { icon: "→", tab: "", desc: "Découvrir la signification de chaque icône obtenue" },
      { icon: "→", tab: "", desc: "Voir combien de fois vous l'avez gagnée (score x/y)" },
      { icon: "→", tab: "", desc: "Comprendre les grands enjeux éthiques associés (cartes A à H)" },
    ],
    visual: "icons",
  },
  {
    emoji: "📊",
    title: "Comparez avec les autres !",
    color: "#ce93d8",
    text: "Dans l'onglet Liste, les pourcentages et le nombre de votes s'affichent directement sous chaque question, pour comparer votre choix avec l'ensemble des participants.",
    sub: "Les résultats n'apparaissent qu'après votre propre vote sur chaque question.",
    visual: "compare",
  },
  {
    emoji: "🔒",
    title: "Vos données restent anonymes",
    color: "#69f0ae",
    text: "Vos choix et votre profil ne sont utilisés qu'à des fins statistiques, pour calculer les pourcentages de réponses affichés dans l'application.",
    sub: "Aucune information personnelle n'est collectée ni traitée : pas de nom, pas d'email, rien qui permette de vous identifier.",
    visual: "privacy",
  },
];

function TutorialModal({ onClose }) {
  const [step, setStep] = useState(0);
  const total = TUTORIAL_STEPS.length;
  const s = TUTORIAL_STEPS[step];

  const renderVisual = () => {
    if (s.visual === "welcome") return (
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 16, padding: "16px 0" }}>
        <div style={{ fontSize: 48 }}>🤖</div>
        <div style={{ fontSize: 32, color: "#ffd764" }}>⚖️</div>
        <div style={{ fontSize: 48 }}>🧠</div>
      </div>
    );
    if (s.visual === "ethics") return (
      <div style={{ display: "flex", justifyContent: "center", gap: 16, padding: "12px 0" }}>
        {[
          { label: "La réponse attendue", icon: "🚫", bg: "rgba(255,82,82,0.1)", border: "rgba(255,82,82,0.3)", color: "#ff8a80" },
          { label: "Ce que vous ressentez", icon: "💡", bg: "rgba(206,147,216,0.12)", border: "rgba(206,147,216,0.4)", color: "#ce93d8" },
        ].map((b, i) => (
          <div key={i} style={{ flex: 1, maxWidth: 140, padding: "12px 10px", background: b.bg, border: "2px solid " + b.border, borderRadius: 12, textAlign: "center" }}>
            <div style={{ fontSize: 28, marginBottom: 6 }}>{b.icon}</div>
            <p style={{ margin: 0, fontSize: 12, fontWeight: "bold", color: b.color }}>{b.label}</p>
          </div>
        ))}
      </div>
    );
    if (s.visual === "cards") return (
      <div style={{ display: "flex", justifyContent: "center", gap: 10, padding: "12px 0" }}>
        {["Oui, l'IA peut m'aider...", "Non, je préfère..."].map((txt, i) => (
          <div key={i} style={{ flex: 1, maxWidth: 150, padding: "10px 12px", borderRadius: 12, border: "2px solid " + (i === 0 ? "rgba(79,195,247,0.5)" : "rgba(206,147,216,0.5)"), background: i === 0 ? "rgba(79,195,247,0.08)" : "rgba(206,147,216,0.08)" }}>
            <div style={{ fontSize: 9, fontWeight: "bold", color: i === 0 ? "#4fc3f7" : "#ce93d8", marginBottom: 4 }}>OPTION {i + 1}</div>
            <div style={{ fontSize: 11, color: "#c0b8a8", lineHeight: 1.4 }}>{txt}</div>
          </div>
        ))}
      </div>
    );
    if (s.visual === "icons") return (
      <div style={{ display: "flex", justifyContent: "center", gap: 12, padding: "14px 0 4px" }}>
        {["img_bonpublic.png","img_Integrite.png","img_Frugalite.png","img_USERascendant.png"].map((img, i) => (
          <div key={i} style={{ position: "relative" }}>
            <img src={ICON_DATA[img]} alt="" style={{ width: 44, height: 44, objectFit: "contain", filter: i < 2 ? "none" : "grayscale(100%) opacity(0.25)" }} />
            <div style={{ position: "absolute", top: -4, right: -8, background: i < 2 ? "#ffd764" : "rgba(255,255,255,0.1)", color: i < 2 ? "#0f0c29" : "#665e52", borderRadius: 8, padding: "1px 5px", fontSize: 9, fontWeight: 900 }}>
              {i < 2 ? (i === 0 ? "3/7" : "2/9") : "?"}
            </div>
          </div>
        ))}
      </div>
    );
    if (s.visual === "compare") return (
      <div style={{ background: "rgba(206,147,216,0.08)", border: "1px solid rgba(206,147,216,0.2)", borderRadius: 10, padding: "12px 14px", margin: "12px 0" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
          <span style={{ fontSize: 11, color: "#4fc3f7", fontWeight: "bold" }}>43%</span>
          <div style={{ flex: 1, height: 6, background: "rgba(255,255,255,0.08)", borderRadius: 3, overflow: "hidden" }}>
            <div style={{ width: "43%", height: "100%", background: "#4fc3f7", display: "inline-block" }} />
            <div style={{ width: "57%", height: "100%", background: "#ce93d8", display: "inline-block" }} />
          </div>
          <span style={{ fontSize: 11, color: "#ce93d8", fontWeight: "bold" }}>57%</span>
        </div>
        <p style={{ margin: 0, fontSize: 11, color: "#665e52", fontStyle: "italic" }}>Exemple de répartition après plusieurs participants</p>
      </div>
    );
    if (s.visual === "privacy") return (
      <div style={{ display: "flex", justifyContent: "center", gap: 14, padding: "10px 0" }}>
        {[
          { icon: "🚫", label: "Pas de nom" },
          { icon: "🚫", label: "Pas d'email" },
          { icon: "✅", label: "Choix anonymisés" },
        ].map((b, i) => (
          <div key={i} style={{ flex: 1, maxWidth: 110, padding: "10px 8px", background: "rgba(105,240,174,0.07)", border: "1px solid rgba(105,240,174,0.25)", borderRadius: 12, textAlign: "center" }}>
            <div style={{ fontSize: 22, marginBottom: 4 }}>{b.icon}</div>
            <p style={{ margin: 0, fontSize: 10.5, fontWeight: "bold", color: "#69f0ae" }}>{b.label}</p>
          </div>
        ))}
      </div>
    );
    return null;
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.82)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, padding: 16, backdropFilter: "blur(6px)" }}>
      <div style={{ maxWidth: 480, width: "100%", background: "linear-gradient(135deg, #1a1735, #2a2550)", border: "2px solid " + s.color + "44", borderRadius: 20, overflow: "hidden", boxShadow: "0 0 60px " + s.color + "22" }}>

        {/* Progress bar */}
        <div style={{ height: 4, background: "rgba(255,255,255,0.06)" }}>
          <div style={{ height: "100%", width: ((step+1)/total*100) + "%", background: s.color, transition: "width 0.4s ease" }} />
        </div>

        {/* Header */}
        <div style={{ padding: "20px 24px 0", display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 42, height: 42, borderRadius: "50%", background: s.color + "22", border: "2px solid " + s.color + "44", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, flexShrink: 0 }}>
            {s.emoji}
          </div>
          <div>
            <p style={{ margin: 0, fontSize: 10, color: "#665e52" }}>{step+1} / {total}</p>
            <h2 style={{ margin: 0, fontFamily: "serif", fontSize: 17, fontWeight: 900, color: s.color }}>{s.title}</h2>
          </div>
          <button onClick={onClose} style={{ marginLeft: "auto", background: "none", border: "none", color: "#665e52", fontSize: 20, cursor: "pointer", flexShrink: 0 }}>✕</button>
        </div>

        {/* Content */}
        <div style={{ padding: "16px 24px 4px" }}>
          {s.quote && (
            <div style={{ background: "rgba(206,147,216,0.12)", border: "2px solid rgba(206,147,216,0.3)", borderRadius: 12, padding: "12px 14px", marginBottom: 12 }}>
              <p style={{ fontSize: 13, fontStyle: "italic", color: "#e8e0d0", margin: 0, lineHeight: 1.6 }}>
                {s.quote}
              </p>
            </div>
          )}
          {renderVisual()}
          <p style={{ fontSize: 13, lineHeight: 1.7, color: "#e8e0d0", margin: "8px 0" }}>{s.text}</p>
          {s.sub && <p style={{ fontSize: 12, color: "#a09888", margin: "4px 0 0", lineHeight: 1.5 }}>{s.sub}</p>}
          {s.items && (
            <div style={{ display: "flex", flexDirection: "column", gap: 7, marginTop: 8 }}>
              {s.items.map((item, i) => (
                <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 10, padding: "8px 12px", background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 9 }}>
                  <span style={{ fontSize: 18, flexShrink: 0 }}>{item.icon}</span>
                  <div>
                    {item.tab && <p style={{ margin: "0 0 2px", fontSize: 12, fontWeight: "bold", color: s.color }}>Onglet {item.tab}</p>}
                    <p style={{ margin: 0, fontSize: 11, color: "#a09888", lineHeight: 1.4 }}>{item.desc}</p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{ padding: "12px 24px 20px", display: "flex", gap: 10, alignItems: "center", borderTop: "1px solid rgba(255,255,255,0.06)", marginTop: 8 }}>
          <div style={{ display: "flex", gap: 5, flex: 1 }}>
            {TUTORIAL_STEPS.map((_, i) => (
              <div key={i} onClick={() => setStep(i)} style={{ width: i === step ? 18 : 6, height: 6, borderRadius: 3, background: i === step ? s.color : i < step ? s.color + "66" : "rgba(255,255,255,0.12)", cursor: "pointer", transition: "all 0.3s" }} />
            ))}
          </div>
          <button onClick={onClose} style={{ padding: "7px 16px", borderRadius: 8, border: "1px solid rgba(255,255,255,0.15)", background: "transparent", color: "#665e52", fontSize: 12, cursor: "pointer" }}>
            Passer
          </button>
          {step < total - 1 ? (
            <button onClick={() => setStep(s => s+1)} style={{ padding: "7px 20px", borderRadius: 8, border: "none", background: s.color, color: "#0f0c29", fontSize: 13, fontWeight: "bold", cursor: "pointer" }}>
              Suivant →
            </button>
          ) : (
            <button onClick={onClose} style={{ padding: "7px 20px", borderRadius: 8, border: "none", background: s.color, color: "#0f0c29", fontSize: 13, fontWeight: "bold", cursor: "pointer" }}>
              C'est parti ! 🚀
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
// ── END TUTORIAL ──────────────────────────────────────────────────────

export default function App() {
  const [counts, setCounts] = useState({});
  const [view, setView] = useState("cards");
  const [status, setStatus] = useState("loading");
  const [errorMsg, setErrorMsg] = useState("");
  const [isAdmin, setIsAdmin] = useState(false);
  const [showLoginModal, setShowLoginModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState("");
  const [loginError, setLoginError] = useState("");
  const [activeCard, setActiveCard] = useState(null);
  const [showProfileModal, setShowProfileModal] = useState(false);
  const [selectedParcours, setSelectedParcours] = useState(null);
  const [showTutorial, setShowTutorial] = useState(() => {
    try { return !sessionStorage.getItem("ethiqia_tutorial_seen"); } catch { return true; }
  });
  const [profile, setProfile] = useState(() => sessionStorage.getItem("ethiqia_profile") || "");
  const [voted, setVoted] = useState(() => {
    try { return JSON.parse(sessionStorage.getItem("ethiqia_voted") || "{}"); } catch { return {}; }
  });

  // Load jsPDF from CDN
  useEffect(() => {
    if (!window.jspdf) {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
      document.head.appendChild(s);
    }
  }, []);

  const closeTutorial = () => {
    setShowTutorial(false);
    try { sessionStorage.setItem("ethiqia_tutorial_seen", "1"); } catch {}
  };

  const load = useCallback(async () => {
    setStatus("loading");
    try { setCounts(await fetchCounts()); setStatus("ready"); }
    catch (e) { setStatus("error"); setErrorMsg(e.message); }
  }, []);
  useEffect(() => { load(); }, [load]);

  const saveVoted = (updated) => {
    setVoted(updated);
    try { sessionStorage.setItem("ethiqia_voted", JSON.stringify(updated)); } catch {}
  };

  const handleChoice = async (id, option) => {
    // Vote définitif : une fois voté sur une carte, plus de retour en
    // arrière possible (ni pour l'utilisateur, ni pour l'admin — voir
    // note ci-dessous sur les corrections).
    if (voted[id]) return;
    const prev = counts[id] || emptyCount();
    const optimistic = { ...counts, [id]: { ...prev, [option]: prev[option] + 1 } };
    setCounts(optimistic);
    saveVoted({ ...voted, [id]: option });
    setStatus("saving");
    try { await submitVote(id, option); setStatus("ready"); setActiveCard(null); }
    catch { setCounts(counts); setStatus("error"); setErrorMsg("Impossible de sauvegarder."); }
  };

  // NOTE : la correction manuelle d'un vote (bouton -1 admin) n'a plus
  // d'équivalent sûr avec cette architecture. L'ancien mécanisme réécrivait
  // un compteur partagé ; avec Grist, chaque vote est une ligne indépendante
  // et la donnée n'est accessible en écriture publique qu'en création, pas
  // en suppression (par choix de sécurité — ouvrir la suppression au public
  // exposerait tous les votes, pas seulement les vôtres). Une correction se
  // fait désormais directement dans l'interface Grist, connecté en tant que
  // propriétaire du document.
  const decrement = async () => {
    window.alert("Les corrections de votes se font désormais directement dans Grist (accès propriétaire), plus depuis cette interface.");
  };

  // Même remarque que pour decrement : la remise à zéro globale supposerait
  // de supprimer toutes les lignes de vote, une opération volontairement
  // réservée au propriétaire du document Grist, pas à l'interface publique.
  const clearAll = async () => {
    window.alert("La remise à zéro se fait désormais directement dans Grist (supprimer les lignes de la table de votes), plus depuis cette interface.");
  };

  const exportXLSX = () => {
    const rows = questions.map(q => {
      const c = counts[q.id] || emptyCount(); const tot = c.A + c.B;
      return { "N°": q.id, "Titre": q.titre, "Situation": q.situation, "Option 1": q.altA, "Option 2": q.altB,
        "Votes Opt.1": c.A, "Votes Opt.2": c.B, "Total": tot,
        "% Opt.1": tot > 0 ? Math.round((c.A / tot) * 100) + "%" : "-",
        "% Opt.2": tot > 0 ? Math.round((c.B / tot) * 100) + "%" : "-" };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 4 }, { wch: 28 }, { wch: 55 }, { wch: 55 }, { wch: 55 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 10 }, { wch: 10 }];
    const totalA = Object.values(counts).reduce((s, c) => s + (c.A || 0), 0);
    const totalB = Object.values(counts).reduce((s, c) => s + (c.B || 0), 0);
    const ws2 = XLSX.utils.json_to_sheet([
      { "": "Total Option 1", "Valeur": totalA }, { "": "Total Option 2", "Valeur": totalB }, { "": "Total général", "Valeur": totalA + totalB }
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Résultats");
    XLSX.utils.book_append_sheet(wb, ws2, "Synthèse");
    XLSX.writeFile(wb, "ethiqia_résultats.xlsx");
  };

  const handleLogin = () => {
    if (passwordInput === ADMIN_PASSWORD) { setIsAdmin(true); setShowLoginModal(false); setPasswordInput(""); setLoginError(""); }
    else setLoginError("Mot de passe incorrect.");
  };

  const handleProfileSelect = (val) => {
    setProfile(val); sessionStorage.setItem("ethiqia_profile", val); setShowProfileModal(false);
  };

  const totalA = Object.values(counts).reduce((s, c) => s + (c.A || 0), 0);
  const totalB = Object.values(counts).reduce((s, c) => s + (c.B || 0), 0);
  const totalVotes = totalA + totalB;
  const profileLabel = PROFILES.find(p => p.value === profile)?.label;
  const statusDot = { loading: "#ffd764", saving: "#4fc3f7", ready: "#69f0ae", error: "#ff5252" };
  const statusLabel = { loading: "Chargement…", saving: "Sauvegarde…", ready: "Synchronisé ✓", error: "Erreur" };
  const activeQ = questions.find(q => q.id === activeCard);

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%)", fontFamily: "'Georgia', serif", color: "#e8e0d0" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=Source+Serif+4:ital,wght@0,300;0,400;1,300&display=swap');
        * { box-sizing: border-box; }
        .tab { padding: 7px 15px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.2); background: transparent; color: #b0a898; cursor: pointer; font-family: 'Source Serif 4', serif; font-size: 12px; transition: all 0.2s; white-space: nowrap; }
        .tab.active { background: rgba(255,215,100,0.15); border-color: rgba(255,215,100,0.5); color: #ffd764; }
        .btn-action { padding: 8px 16px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.2); background: rgba(255,255,255,0.07); color: #e8e0d0; cursor: pointer; font-family: 'Source Serif 4', serif; font-size: 12px; transition: all 0.2s; white-space: nowrap; }
        .btn-action:hover { background: rgba(255,255,255,0.14); }
        .num-card { width: 58px; height: 58px; border-radius: 12px; border: 2px solid rgba(255,255,255,0.12); background: rgba(255,255,255,0.05); cursor: pointer; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px; transition: all 0.2s; font-family: 'Playfair Display', serif; }
        .num-card:hover { border-color: rgba(255,215,100,0.5); background: rgba(255,215,100,0.08); transform: translateY(-2px); }
        .num-card.voted-a { border-color: rgba(79,195,247,0.6); background: rgba(79,195,247,0.1); }
        .num-card.voted-b { border-color: rgba(206,147,216,0.6); background: rgba(206,147,216,0.1); }
        .modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.72); display: flex; align-items: center; justify-content: center; z-index: 100; padding: 16px; backdrop-filter: blur(4px); }
        .modal { background: linear-gradient(135deg, #1a1735, #2a2550); border: 1px solid rgba(255,255,255,0.15); border-radius: 20px; max-width: 560px; width: 100%; padding: 28px; position: relative; max-height: 90vh; overflow-y: auto; }
        .btn-vote { width: 100%; text-align: left; padding: 13px 16px; border-radius: 12px; border: 2px solid rgba(255,255,255,0.1); cursor: pointer; font-family: 'Source Serif 4', serif; font-size: 13.5px; line-height: 1.5; transition: all 0.2s; background: rgba(255,255,255,0.05); color: #e8e0d0; }
        .btn-vote.opt-a:hover:not(:disabled) { border-color: #4fc3f7; background: rgba(79,195,247,0.12); }
        .btn-vote.opt-b:hover:not(:disabled) { border-color: #ce93d8; background: rgba(206,147,216,0.12); }
        .btn-vote:disabled { opacity: 0.45; cursor: not-allowed; }
        .btn-vote.already-a { border-color: rgba(79,195,247,0.5); background: rgba(79,195,247,0.1); }
        .btn-vote.already-b { border-color: rgba(206,147,216,0.5); background: rgba(206,147,216,0.1); }
        .list-card { background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.12); border-radius: 16px; padding: 20px; margin-bottom: 14px; }
        .vote-row { display: flex; align-items: center; gap: 10px; }
        .count-badge { min-width: 34px; height: 34px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 13px; font-weight: bold; flex-shrink: 0; border: none; }
        .count-a { background: rgba(79,195,247,0.25); color: #4fc3f7; cursor: pointer; }
        .count-b { background: rgba(206,147,216,0.25); color: #ce93d8; cursor: pointer; }
        .count-zero { background: rgba(255,255,255,0.06); color: #665e52; cursor: default; }
        .profile-option { width: 100%; padding: 14px 18px; border-radius: 12px; border: 2px solid rgba(255,255,255,0.1); background: rgba(255,255,255,0.05); color: #e8e0d0; cursor: pointer; font-family: 'Source Serif 4', serif; font-size: 14px; text-align: left; transition: all 0.2s; }
        .profile-option:hover { border-color: #ffd764; background: rgba(255,215,100,0.08); }
        .profile-option.selected { border-color: #ffd764; background: rgba(255,215,100,0.12); color: #ffd764; }
        .summary-row { display: flex; align-items: flex-start; gap: 12px; padding: 11px 0; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .spinner { width: 13px; height: 13px; border: 2px solid rgba(255,255,255,0.1); border-top-color: #4fc3f7; border-radius: 50%; animation: spin 0.7s linear infinite; display: inline-block; }
        @keyframes spin { to { transform: rotate(360deg); } }
      `}</style>

      <div style={{ maxWidth: 820, margin: "0 auto", padding: "24px 14px" }}>

        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 20 }}>
          <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: "clamp(26px,5vw,42px)", fontWeight: 900, margin: 0, background: "linear-gradient(135deg, #ffd764, #ce93d8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Éthiq·IA</h1>
          <p style={{ fontFamily: "'Source Serif 4', serif", fontStyle: "italic", color: "#a09888", fontSize: 12, marginTop: 4 }}>
            {isAdmin ? "Mode administrateur" : "Cliquez sur une carte pour voter"} <button onClick={() => setShowTutorial(true)} style={{ marginLeft: 8, background: "rgba(255,215,100,0.15)", border: "1px solid rgba(255,215,100,0.3)", borderRadius: "50%", width: 22, height: 22, color: "#ffd764", fontSize: 12, cursor: "pointer", verticalAlign: "middle", lineHeight: "20px" }}>?</button>
          </p>
        </div>

        {/* Profile + Stats bar */}
        <div style={{ marginBottom: 16, padding: "12px 18px", background: "rgba(255,255,255,0.05)", borderRadius: 12, border: "1px solid rgba(255,255,255,0.1)", display: "flex", gap: 16, flexWrap: "wrap", alignItems: "center" }}>
          <button className="btn-action" style={{ borderColor: profile ? "rgba(255,215,100,0.4)" : "rgba(255,255,255,0.2)", color: profile ? "#ffd764" : "#b0a898" }} onClick={() => setShowProfileModal(true)}>
            👤 {profileLabel || "Mon profil"}
          </button>
          <div style={{ width: 1, height: 28, background: "rgba(255,255,255,0.1)" }} />
          <div><span style={{ fontSize: 11, color: "#a09888" }}>Total</span><p style={{ margin: "1px 0 0", fontSize: 18, fontWeight: "bold", color: "#ffd764" }}>{totalVotes}</p></div>
          <div><span style={{ fontSize: 11, color: "#a09888" }}>Option 1</span><p style={{ margin: "1px 0 0", fontSize: 18, fontWeight: "bold", color: "#4fc3f7" }}>{totalA}</p></div>
          <div><span style={{ fontSize: 11, color: "#a09888" }}>Option 2</span><p style={{ margin: "1px 0 0", fontSize: 18, fontWeight: "bold", color: "#ce93d8" }}>{totalB}</p></div>
          <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 7 }}>
            {(status === "loading" || status === "saving") && <span className="spinner" />}
            <span style={{ width: 7, height: 7, borderRadius: "50%", background: statusDot[status], display: "inline-block" }} />
            <span style={{ fontSize: 11, color: status === "error" ? "#ff5252" : "#a09888" }}>{statusLabel[status]}</span>
          </div>
        </div>

        {/* Nav */}
        <div style={{ display: "flex", gap: 6, marginBottom: 18, flexWrap: "wrap", alignItems: "center" }}>
          <button className={`tab ${view === "parcours" ? "active" : ""}`} onClick={() => setView("parcours")}>🗺️ Parcours</button>
          <button className={`tab ${view === "cards" ? "active" : ""}`} onClick={() => setView("cards")}>🃏 Cartes</button>
          <button className={`tab ${view === "list" ? "active" : ""}`} onClick={() => setView("list")}>📋 Liste</button>
          <button className={`tab ${view === "bilan" ? "active" : ""}`} onClick={() => setView("bilan")}>🏅 Bilan</button>
          <div style={{ marginLeft: "auto", display: "flex", gap: 6, flexWrap: "wrap" }}>
            <button className="btn-action" onClick={load} title="Actualiser">↺</button>
            <button className="btn-action" style={{ background: "rgba(255,107,107,0.15)", borderColor: "rgba(255,107,107,0.4)", color: "#ff8a80" }} onClick={() => generatePDF(counts, voted, profile)}>
              📄 PDF
            </button>
            {isAdmin && (
              <>
                <button className="btn-action" style={{ background: "rgba(79,195,247,0.15)", borderColor: "rgba(79,195,247,0.4)", color: "#4fc3f7" }} onClick={exportXLSX}>⬇ .xlsx</button>
                <button className="btn-action" style={{ color: "#ff8a80" }} onClick={clearAll}>✕ Réinitialiser</button>
                <button className="btn-action" style={{ color: "#69f0ae", borderColor: "rgba(105,240,174,0.3)" }} onClick={() => setIsAdmin(false)}>Déco.</button>
              </>
            )}
            {!isAdmin && <button className="btn-action" onClick={() => setShowLoginModal(true)}>🔒 Admin</button>}
          </div>
        </div>

        {status === "loading" && (
          <div style={{ textAlign: "center", padding: "60px 0", color: "#a09888" }}>
            <span className="spinner" style={{ width: 28, height: 28, borderWidth: 3 }} />
            <p style={{ marginTop: 16 }}>Connexion à Supabase…</p>
          </div>
        )}


        {/* PARCOURS */}
        {status !== "loading" && view === "parcours" && (
          <div>
            <p style={{ fontSize: 13, color: "#a09888", marginBottom: 18, fontStyle: "italic" }}>
              Sélectionnez un parcours pour mettre en surbrillance les cartes correspondantes dans la vue Cartes.
            </p>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {PARCOURS.map(p => {
                const isSelected = selectedParcours === p.id;
                const votedInParcours = p.cards.filter(id => voted[id]).length;
                return (
                  <div key={p.id} onClick={() => { setSelectedParcours(isSelected ? null : p.id); setView("cards"); }}
                    style={{
                      padding: "14px 18px", borderRadius: 14, cursor: "pointer", transition: "all 0.2s",
                      border: "2px solid " + (isSelected ? p.color : "rgba(255,255,255,0.1)"),
                      background: isSelected ? p.color + "18" : "rgba(255,255,255,0.04)",
                      boxShadow: isSelected ? "0 0 16px " + p.color + "44" : "none",
                    }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                      {/* Color dot */}
                      <div style={{ width: 14, height: 14, borderRadius: "50%", background: p.color, flexShrink: 0, boxShadow: "0 0 8px " + p.color }} />
                      {/* Label */}
                      <div style={{ flex: 1 }}>
                        <p style={{ margin: 0, fontSize: 14, fontWeight: "bold", color: isSelected ? p.color : "#f5efe0", fontFamily: "'Playfair Display', serif" }}>{p.label}</p>
                        <p style={{ margin: "3px 0 0", fontSize: 11, color: "#665e52" }}>
                          {p.cards.length} cartes · {votedInParcours} votée{votedInParcours > 1 ? "s" : ""}
                        </p>
                      </div>
                      {/* Card numbers */}
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 4, justifyContent: "flex-end", maxWidth: 160 }}>
                        {p.cards.map(id => (
                          <span key={id} style={{
                            width: 24, height: 24, borderRadius: 6, fontSize: 10, fontWeight: "bold",
                            display: "flex", alignItems: "center", justifyContent: "center",
                            background: voted[id] ? p.color + "33" : "rgba(255,255,255,0.06)",
                            color: voted[id] ? p.color : "#665e52",
                            border: "1px solid " + (voted[id] ? p.color + "55" : "rgba(255,255,255,0.08)"),
                          }}>{id}</span>
                        ))}
                      </div>
                      {/* Arrow */}
                      <span style={{ color: isSelected ? p.color : "#443d36", fontSize: 16, flexShrink: 0 }}>
                        {isSelected ? "✓" : "→"}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
            {selectedParcours && (
              <button onClick={() => setSelectedParcours(null)}
                style={{ marginTop: 14, padding: "8px 18px", borderRadius: 10, border: "1px solid rgba(255,255,255,0.15)", background: "transparent", color: "#a09888", fontSize: 12, cursor: "pointer" }}>
                Effacer la sélection
              </button>
            )}
          </div>
        )}

        {/* CARDS */}
        {status !== "loading" && view === "cards" && (
          <>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 9 }}>
              {questions.map(q => {
                const c = counts[q.id] || emptyCount();
                const tot = c.A + c.B;
                const uv = voted[q.id];
                const parcours = selectedParcours ? PARCOURS.find(p => p.id === selectedParcours) : null;
                const inParcours = parcours && parcours.cards.includes(q.id);
                const pColor = inParcours ? parcours.color : null;
                return (
                  <button key={q.id} className={`num-card${uv === "A" ? " voted-a" : uv === "B" ? " voted-b" : ""}`} onClick={() => setActiveCard(q.id)} title={q.titre}
                    style={{ height: uv ? "auto" : undefined, minHeight: 58, padding: uv ? "6px 4px" : undefined,
                      borderColor: pColor || undefined,
                      background: pColor ? pColor + "22" : undefined,
                      boxShadow: pColor ? "0 0 8px " + pColor + "55" : undefined,
                    }}>
                    <span style={{ fontSize: 17, fontWeight: 900, color: uv === "A" ? "#4fc3f7" : uv === "B" ? "#ce93d8" : "#f5efe0" }}>{q.id}</span>
                    {uv && getIcons(q.id, uv).length > 0 && (
                      <div style={{ display: "flex", gap: 2, justifyContent: "center", flexWrap: "wrap" }}>
                        {getIcons(q.id, uv).map((src, i) => (
                          <img key={i} src={src} alt="" style={{ width: 22, height: 22, objectFit: "contain", opacity: 0.9 }} />
                        ))}
                      </div>
                    )}
                    {tot > 0 && <span style={{ fontSize: 7, color: "#a09888" }}>{tot}v</span>}
                  </button>
                );
              })}
            </div>
            <p style={{ fontSize: 10, color: "#554e46", marginTop: 14, textAlign: "center" }}>
              Bleu = option 1 · Violet = option 2 {!isAdmin && "· Un seul vote par carte"}
            </p>
          </>
        )}

        {/* LIST */}
        {status !== "loading" && view === "list" && questions.map(q => {
          const c = counts[q.id] || emptyCount(); const tot = c.A + c.B;
          const pctA = tot > 0 ? Math.round((c.A / tot) * 100) : null;
          const uv = voted[q.id]; const alreadyVoted = !isAdmin && !!uv;
          return (
            <div key={q.id} className="list-card">
              <div style={{ display: "flex", alignItems: "flex-start", gap: 10, marginBottom: 8 }}>
                <div style={{ width: 24, height: 24, borderRadius: "50%", background: "rgba(255,255,255,0.07)", color: "#a09888", fontSize: 10, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginTop: 2 }}>{q.id}</div>
                <div style={{ flex: 1 }}>
                  <h3 style={{ margin: "0 0 3px", fontFamily: "'Playfair Display', serif", fontSize: 14, fontWeight: 700, color: "#f5efe0" }}>{q.titre}</h3>
                  <p style={{ margin: 0, fontSize: 11.5, color: "#a09888", fontStyle: "italic", lineHeight: 1.45 }}>{q.situation}</p>
                </div>
                {uv && tot > 0 && <span style={{ fontSize: 10, color: "#665e52", whiteSpace: "nowrap" }}>{tot}v</span>}
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                <div className="vote-row">
                  <button className={`btn-vote opt-a${uv === "A" ? " already-a" : ""}`} disabled={alreadyVoted && uv !== "A"} onClick={() => handleChoice(q.id, "A")}>
                    <span style={{ fontSize: 9, fontWeight: "bold", color: "#4fc3f7", display: "block", marginBottom: 2 }}>OPTION 1{uv === "A" ? " ✓" : ""}</span>{q.altA}
                  </button>
                  {uv === "A" && getIcons(q.id, "A").length > 0 && (
                    <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                      {getIcons(q.id, "A").map((src, i) => (
                        <img key={i} src={src} alt="" style={{ width: 38, height: 38, objectFit: "contain", borderRadius: 6, background: "rgba(79,195,247,0.1)", padding: 3 }} />
                      ))}
                    </div>
                  )}
                  {isAdmin && <button className={`count-badge ${c.A > 0 ? "count-a" : "count-zero"}`} onClick={() => decrement(q.id, "A")}>{c.A}</button>}
                </div>
                <div className="vote-row">
                  <button className={`btn-vote opt-b${uv === "B" ? " already-b" : ""}`} disabled={alreadyVoted && uv !== "B"} onClick={() => handleChoice(q.id, "B")}>
                    <span style={{ fontSize: 9, fontWeight: "bold", color: "#ce93d8", display: "block", marginBottom: 2 }}>OPTION 2{uv === "B" ? " ✓" : ""}</span>{q.altB}
                  </button>
                  {uv === "B" && getIcons(q.id, "B").length > 0 && (
                    <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                      {getIcons(q.id, "B").map((src, i) => (
                        <img key={i} src={src} alt="" style={{ width: 38, height: 38, objectFit: "contain", borderRadius: 6, background: "rgba(206,147,216,0.1)", padding: 3 }} />
                      ))}
                    </div>
                  )}
                  {isAdmin && <button className={`count-badge ${c.B > 0 ? "count-b" : "count-zero"}`} onClick={() => decrement(q.id, "B")}>{c.B}</button>}
                </div>
                <div style={{ marginTop: 4, padding: "8px 12px", background: "rgba(255,255,255,0.04)", borderRadius: 10, border: "1px solid rgba(255,255,255,0.08)" }}>
                  {uv && tot > 0 && pctA !== null ? (
                    <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 12, color: "#4fc3f7", fontWeight: "bold" }}>{pctA}%</span>
                      <span style={{ fontSize: 11, color: "#665e52" }}>Opt.1</span>
                      <span style={{ fontSize: 11, color: "#554e46" }}>·</span>
                      <span style={{ fontSize: 12, color: "#ce93d8", fontWeight: "bold" }}>{100 - pctA}%</span>
                      <span style={{ fontSize: 11, color: "#665e52" }}>Opt.2</span>
                      <span style={{ fontSize: 11, color: "#443d36", marginLeft: 4 }}>({tot} vote{tot > 1 ? "s" : ""})</span>
                      <div style={{ flex: 1, minWidth: 80, height: 4, background: "rgba(255,255,255,0.08)", borderRadius: 2, overflow: "hidden" }}>
                        <div style={{ width: `${pctA}%`, height: "100%", background: "#4fc3f7", display: "inline-block" }} />
                        <div style={{ width: `${100 - pctA}%`, height: "100%", background: "#ce93d8", display: "inline-block" }} />
                      </div>
                    </div>
                  ) : (
                    <p style={{ margin: 0, fontSize: 10, color: "#443d36", fontStyle: "italic", textAlign: "center" }}>
                      {uv ? "Aucun vote enregistré pour cette carte." : "Votez pour voir les résultats et les icônes."}
                    </p>
                  )}
                </div>
              </div>
            </div>
          );
        })}

        {/* SUMMARY */}
        {status !== "loading" && view === "summary_disabled" && questions.map(q => {
          const c = counts[q.id] || emptyCount(); const tot = c.A + c.B;
          const pctA = tot > 0 ? Math.round((c.A / tot) * 100) : null;
          const uv = voted[q.id];
          const alreadyVoted = !isAdmin && !!uv;
          const iconsA = getIcons(q.id, "A");
          const iconsB = getIcons(q.id, "B");
          return (
            <div key={q.id} className="list-card">
              {/* Header: numéro + titre + situation */}
              <div style={{ display: "flex", alignItems: "flex-start", gap: 10, marginBottom: 8 }}>
                <div style={{ width: 24, height: 24, borderRadius: "50%", background: "rgba(255,255,255,0.07)", color: "#a09888", fontSize: 10, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginTop: 2 }}>{q.id}</div>
                <div style={{ flex: 1 }}>
                  <h3 style={{ margin: "0 0 3px", fontFamily: "'Playfair Display', serif", fontSize: 14, fontWeight: 700, color: "#f5efe0" }}>{q.titre}</h3>
                  <p style={{ margin: 0, fontSize: 11.5, color: "#a09888", fontStyle: "italic", lineHeight: 1.45 }}>{q.situation}</p>
                </div>
                {tot > 0 && <span style={{ fontSize: 10, color: "#665e52", whiteSpace: "nowrap" }}>{tot}v</span>}
              </div>

              {/* Options */}
              <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>

                {/* Option A */}
                <div className="vote-row">
                  <button className={`btn-vote opt-a${uv === "A" ? " already-a" : ""}`}
                    disabled={alreadyVoted && uv !== "A"}
                    onClick={() => handleChoice(q.id, "A")}>
                    <span style={{ fontSize: 9, fontWeight: "bold", color: "#4fc3f7", display: "block", marginBottom: 2 }}>
                      OPTION 1{uv === "A" ? " ✓" : ""}
                    </span>
                    {q.altA}
                  </button>
                  {uv === "A" && iconsA.length > 0 && (
                    <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                      {iconsA.map((src, i) => (
                        <img key={i} src={src} alt="" style={{ width: 38, height: 38, objectFit: "contain", borderRadius: 6, background: "rgba(79,195,247,0.1)", padding: 3 }} />
                      ))}
                    </div>
                  )}
                  {isAdmin && <button className={`count-badge ${c.A > 0 ? "count-a" : "count-zero"}`} onClick={() => decrement(q.id, "A")}>{c.A}</button>}
                </div>

                {/* Option B */}
                <div className="vote-row">
                  <button className={`btn-vote opt-b${uv === "B" ? " already-b" : ""}`}
                    disabled={alreadyVoted && uv !== "B"}
                    onClick={() => handleChoice(q.id, "B")}>
                    <span style={{ fontSize: 9, fontWeight: "bold", color: "#ce93d8", display: "block", marginBottom: 2 }}>
                      OPTION 2{uv === "B" ? " ✓" : ""}
                    </span>
                    {q.altB}
                  </button>
                  {uv === "B" && iconsB.length > 0 && (
                    <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                      {iconsB.map((src, i) => (
                        <img key={i} src={src} alt="" style={{ width: 38, height: 38, objectFit: "contain", borderRadius: 6, background: "rgba(206,147,216,0.1)", padding: 3 }} />
                      ))}
                    </div>
                  )}
                  {isAdmin && <button className={`count-badge ${c.B > 0 ? "count-b" : "count-zero"}`} onClick={() => decrement(q.id, "B")}>{c.B}</button>}
                </div>

                {/* Feedback votes — visible uniquement après avoir voté */}
                {uv && tot > 0 && pctA !== null && (
                  <div style={{ marginTop: 4, padding: "8px 12px", background: "rgba(255,255,255,0.04)", borderRadius: 10, border: "1px solid rgba(255,255,255,0.08)" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 12, color: "#4fc3f7", fontWeight: "bold" }}>{pctA}%</span>
                      <span style={{ fontSize: 11, color: "#665e52" }}>Opt.1</span>
                      <span style={{ fontSize: 11, color: "#554e46" }}>·</span>
                      <span style={{ fontSize: 12, color: "#ce93d8", fontWeight: "bold" }}>{100 - pctA}%</span>
                      <span style={{ fontSize: 11, color: "#665e52" }}>Opt.2</span>
                      <span style={{ fontSize: 11, color: "#443d36", marginLeft: 4 }}>({tot} vote{tot > 1 ? "s" : ""})</span>
                      <div style={{ flex: 1, minWidth: 80, height: 4, background: "rgba(255,255,255,0.08)", borderRadius: 2, overflow: "hidden" }}>
                        <div style={{ width: `${pctA}%`, height: "100%", background: "#4fc3f7", display: "inline-block" }} />
                        <div style={{ width: `${100 - pctA}%`, height: "100%", background: "#ce93d8", display: "inline-block" }} />
                      </div>
                    </div>
                    {!uv && <p style={{ margin: "6px 0 0", fontSize: 10, color: "#443d36", fontStyle: "italic" }}>Votez pour voir la répartition.</p>}
                  </div>
                )}

                {/* Invitation à voter si pas encore fait */}
                {!uv && (
                  <p style={{ margin: "2px 0 0", fontSize: 10, color: "#443d36", fontStyle: "italic", textAlign: "center" }}>
                    Votez pour voir les résultats et les icônes.
                  </p>
                )}

              </div>
            </div>
          );
        })}
      </div>

        {/* BILAN */}
        {status !== "loading" && view === "bilan" && (() => {
          // Count icon occurrences from user votes
          const iconCount = {};
          Object.entries(voted).forEach(([qid, opt]) => {
            const icons = getIcons(parseInt(qid), opt);
            icons.forEach(uri => {
              // find filename from ICON_DATA
              const fname = Object.keys(ICON_DATA).find(k => ICON_DATA[k] === uri);
              if (fname) iconCount[fname] = (iconCount[fname] || 0) + 1;
            });
          });

          const totalVoted = Object.keys(voted).length;

          return (
            <div>
              {/* Header bilan */}
              <div style={{ marginBottom: 20, padding: "14px 20px", background: "rgba(255,255,255,0.05)", borderRadius: 12, border: "1px solid rgba(255,255,255,0.1)" }}>
                <p style={{ margin: 0, fontSize: 13, color: "#a09888" }}>
                  {totalVoted === 0
                    ? "Vous n'avez pas encore voté. Commencez par répondre aux questions pour voir votre bilan !"
                    : `Bilan basé sur ${totalVoted} vote${totalVoted > 1 ? "s" : ""}. Chaque icône représente une valeur associée à vos choix.`}
                </p>
              </div>

              {/* Cards A-H — enjeux */}
              <div style={{ marginBottom: 24 }}>
                <h3 style={{ fontFamily: "'Playfair Display', serif", fontSize: 16, color: "#ffd764", margin: "0 0 14px" }}>
                  Les enjeux de chaque tension
                </h3>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 12 }}>
                  {CARDS_AH.map((card, idx) => {
                    // Check if any icon of this card has been obtained
                    const scores = card.icons.map(img => iconCount[img] || 0);
                    const hasAny = scores.some(s => s > 0);
                    return (
                      <div key={idx} style={{
                        width: "calc(50% - 6px)", minWidth: 260,
                        background: hasAny ? "linear-gradient(135deg, #2a1f5a, #3d2b7a)" : "rgba(255,255,255,0.03)",
                        border: hasAny ? "2px solid rgba(180,130,255,0.45)" : "1px solid rgba(255,255,255,0.08)",
                        borderRadius: 14, overflow: "hidden",
                      }}>
                        {/* Color bar */}
                        <div style={{ height: 5, background: hasAny ? "linear-gradient(90deg, #9b59b6, #ce93d8)" : "rgba(255,255,255,0.06)" }} />
                        {/* Header */}
                        <div style={{ padding: "12px 14px 0", display: "flex", alignItems: "center", gap: 10 }}>
                          <div style={{
                            width: 28, height: 28, borderRadius: "50%", flexShrink: 0,
                            background: "rgba(255,215,100,0.15)", color: "#ffd764",
                            fontFamily: "'Playfair Display', serif", fontWeight: 900, fontSize: 14,
                            display: "flex", alignItems: "center", justifyContent: "center",
                          }}>{card.level}</div>
                          <p style={{ margin: 0, fontFamily: "'Playfair Display', serif", fontSize: 13, fontWeight: 700, color: "#f5efe0", lineHeight: 1.3 }}>{card.title}</p>
                        </div>
                        {/* Duo image */}
                        <div style={{ padding: "10px 14px 4px", display: "flex", justifyContent: "center" }}>
                          <img src={card.duoImg} alt="" style={{
                            height: 52, objectFit: "contain",
                            filter: hasAny ? "none" : "grayscale(80%) opacity(0.3)",
                            transition: "all 0.3s",
                          }} />
                        </div>
                        {/* Question */}
                        <div style={{ padding: "4px 14px 8px" }}>
                          <p style={{ margin: 0, fontSize: 11, fontStyle: "italic", color: hasAny ? "#c0b8a8" : "#554e46", lineHeight: 1.4 }}>{card.question}</p>
                        </div>
                        {/* Content — collapsible */}
                        {hasAny && (
                          <div style={{ padding: "0 14px 12px" }}>
                            <p style={{ margin: 0, fontSize: 11, color: "#a09888", lineHeight: 1.5 }}>{card.content}</p>
                            <div style={{ marginTop: 8, display: "flex", gap: 6 }}>
                              {card.icons.map((img, i) => (
                                <div key={i} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                                  <img src={BILAN_ICON_DATA[img]} alt="" style={{ width: 18, height: 18, objectFit: "contain", opacity: 0.7 }} />
                                  <span style={{ fontSize: 10, color: i === 0 ? "#4fc3f7" : "#ce93d8" }}>
                                    {iconCount[img] || 0}/{ICON_MAX[img] || "?"}
                                  </span>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}
                        {!hasAny && (
                          <p style={{ margin: "0 14px 12px", fontSize: 10, color: "#443d36", fontStyle: "italic" }}>
                            Votez sur les questions liées pour révéler cet enjeu.
                          </p>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Séparateur */}
              <div style={{ margin: "4px 0 20px", borderTop: "1px solid rgba(255,255,255,0.08)" }} />
              <h3 style={{ fontFamily: "'Playfair Display', serif", fontSize: 16, color: "#ffd764", margin: "0 0 14px" }}>
                Détail de tes icônes
              </h3>

              {/* Cards grid */}
              <div style={{ display: "flex", flexWrap: "wrap", gap: 14, justifyContent: "flex-start" }}>
                {BILAN_CARDS.map((card, idx) => {
                  const count1 = iconCount[card.img1] || 0;
                  const count2 = iconCount[card.img2] || 0;
                  const hasVotes = count1 > 0 || count2 > 0;
                  return (
                    <div key={idx} style={{
                      width: "calc(50% - 7px)", minWidth: 280,
                      background: hasVotes ? "linear-gradient(135deg, #2a1f5a, #3d2b7a)" : "rgba(255,255,255,0.04)",
                      border: hasVotes ? "2px solid rgba(180,130,255,0.5)" : "1px solid rgba(255,255,255,0.1)",
                      borderRadius: 16, overflow: "hidden",
                      transition: "all 0.3s",
                      opacity: totalVoted === 0 ? 0.4 : 1,
                    }}>
                      {/* Card top color bar */}
                      <div style={{ height: 6, background: hasVotes ? "linear-gradient(90deg, #9b59b6, #ce93d8)" : "rgba(255,255,255,0.08)" }} />

                      {/* Two icons side by side */}
                      <div style={{ display: "flex", gap: 0 }}>
                        {/* Icon 1 */}
                        <div style={{
                          flex: 1, padding: "16px 12px 12px",
                          background: count1 > 0 ? "rgba(79,195,247,0.08)" : "transparent",
                          borderRight: "1px solid rgba(255,255,255,0.06)",
                          display: "flex", flexDirection: "column", alignItems: "center", gap: 8,
                        }}>
                          <div style={{ position: "relative" }}>
                            <img src={BILAN_ICON_DATA[card.img1]} alt="" style={{
                              width: 56, height: 56, objectFit: "contain",
                              filter: count1 > 0 ? "none" : "grayscale(100%) opacity(0.3)",
                              transition: "all 0.3s",
                            }} />
                            <div style={{
                              position: "absolute", top: -8, right: -18,
                              background: count1 > 0 ? "#4fc3f7" : "rgba(255,255,255,0.1)",
                              color: count1 > 0 ? "#0f0c29" : "#665e52",
                              borderRadius: 10, padding: "1px 6px",
                              fontSize: 10, fontWeight: 900, whiteSpace: "nowrap",
                            }}>{count1}/{ICON_MAX[card.img1] || "?"}</div>
                          </div>
                          <p style={{ margin: 0, fontSize: 11, fontWeight: "bold", color: count1 > 0 ? "#4fc3f7" : "#554e46", textAlign: "center", lineHeight: 1.3 }}>{card.title1}</p>
                          <p style={{ margin: 0, fontSize: 10, color: count1 > 0 ? "#b0a898" : "#443d36", textAlign: "center", lineHeight: 1.4 }}>{card.desc1}</p>
                        </div>

                        {/* Icon 2 */}
                        <div style={{
                          flex: 1, padding: "16px 12px 12px",
                          background: count2 > 0 ? "rgba(206,147,216,0.08)" : "transparent",
                          display: "flex", flexDirection: "column", alignItems: "center", gap: 8,
                        }}>
                          <div style={{ position: "relative" }}>
                            <img src={BILAN_ICON_DATA[card.img2]} alt="" style={{
                              width: 56, height: 56, objectFit: "contain",
                              filter: count2 > 0 ? "none" : "grayscale(100%) opacity(0.3)",
                              transition: "all 0.3s",
                            }} />
                            <div style={{
                              position: "absolute", top: -8, right: -18,
                              background: count2 > 0 ? "#ce93d8" : "rgba(255,255,255,0.1)",
                              color: count2 > 0 ? "#0f0c29" : "#665e52",
                              borderRadius: 10, padding: "1px 6px",
                              fontSize: 10, fontWeight: 900, whiteSpace: "nowrap",
                            }}>{count2}/{ICON_MAX[card.img2] || "?"}</div>
                          </div>
                          <p style={{ margin: 0, fontSize: 11, fontWeight: "bold", color: count2 > 0 ? "#ce93d8" : "#554e46", textAlign: "center", lineHeight: 1.3 }}>{card.title2}</p>
                          <p style={{ margin: 0, fontSize: 10, color: count2 > 0 ? "#b0a898" : "#443d36", textAlign: "center", lineHeight: 1.4 }}>{card.desc2}</p>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              <p style={{ textAlign: "center", fontSize: 11, color: "#443d36", marginTop: 20 }}>
                Les icônes grisées n'ont pas encore été obtenues. Le compteur indique le nombre de fois où vous avez choisi cette valeur.
              </p>
            </div>
          );
        })()}

      {/* CARD MODAL */}
      {activeCard !== null && activeQ && (() => {
        const c = counts[activeQ.id] || emptyCount(); const tot = c.A + c.B;
        const uv = voted[activeQ.id]; const alreadyVoted = !isAdmin && !!uv;
        return (
          <div className="modal-overlay" onClick={() => setActiveCard(null)}>
            <div className="modal" onClick={e => e.stopPropagation()}>
              <button onClick={() => setActiveCard(null)} style={{ position: "absolute", top: 14, right: 14, background: "none", border: "none", color: "#a09888", fontSize: 18, cursor: "pointer" }}>✕</button>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                <div style={{ width: 30, height: 30, borderRadius: "50%", background: "rgba(255,215,100,0.15)", color: "#ffd764", fontFamily: "'Playfair Display', serif", fontWeight: 900, fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{activeQ.id}</div>
                <h2 style={{ margin: 0, fontFamily: "'Playfair Display', serif", fontSize: 19, fontWeight: 900, color: "#f5efe0" }}>{activeQ.titre}</h2>
              </div>
              <p style={{ fontSize: 13.5, color: "#c0b8a8", lineHeight: 1.6, marginBottom: 18, fontStyle: "italic", borderLeft: "3px solid rgba(255,215,100,0.3)", paddingLeft: 12 }}>{activeQ.situation}</p>
              {alreadyVoted && <p style={{ fontSize: 11, color: "#ffd764", marginBottom: 10, textAlign: "center" }}>✓ Vous avez déjà voté pour cette carte.</p>}
              <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                <button className={`btn-vote opt-a${uv === "A" ? " already-a" : ""}`} disabled={alreadyVoted && uv !== "A"} onClick={() => handleChoice(activeQ.id, "A")}>
                  <span style={{ fontSize: 9, fontWeight: "bold", color: "#4fc3f7", display: "block", marginBottom: 3 }}>OPTION 1{uv === "A" ? " ✓ votre choix" : ""}</span>{activeQ.altA}
                </button>
                <button className={`btn-vote opt-b${uv === "B" ? " already-b" : ""}`} disabled={alreadyVoted && uv !== "B"} onClick={() => handleChoice(activeQ.id, "B")}>
                  <span style={{ fontSize: 9, fontWeight: "bold", color: "#ce93d8", display: "block", marginBottom: 3 }}>OPTION 2{uv === "B" ? " ✓ votre choix" : ""}</span>{activeQ.altB}
                </button>
              </div>
              {uv && tot > 0 && <div style={{ marginTop: 12, display: "flex", gap: 7, alignItems: "center" }}>
                <span style={{ fontSize: 10, color: "#4fc3f7" }}>{c.A}</span>
                <div style={{ flex: 1, height: 4, background: "rgba(255,255,255,0.08)", borderRadius: 2, overflow: "hidden" }}>
                  <div style={{ width: `${(c.A / tot) * 100}%`, height: "100%", background: "#4fc3f7", display: "inline-block" }} />
                  <div style={{ width: `${(c.B / tot) * 100}%`, height: "100%", background: "#ce93d8", display: "inline-block" }} />
                </div>
                <span style={{ fontSize: 10, color: "#ce93d8" }}>{c.B}</span>
              </div>}
            </div>
          </div>
        );
      })()}

      {/* PROFILE MODAL */}
      {showProfileModal && (
        <div className="modal-overlay" onClick={() => setShowProfileModal(false)}>
          <div className="modal" style={{ maxWidth: 400 }} onClick={e => e.stopPropagation()}>
            <button onClick={() => setShowProfileModal(false)} style={{ position: "absolute", top: 14, right: 14, background: "none", border: "none", color: "#a09888", fontSize: 18, cursor: "pointer" }}>✕</button>
            <h2 style={{ margin: "0 0 6px", fontFamily: "'Playfair Display', serif", fontSize: 19, color: "#f5efe0" }}>Mon profil</h2>
            <p style={{ margin: "0 0 18px", fontSize: 12, color: "#a09888", fontStyle: "italic" }}>Sélectionnez votre niveau d'études</p>
            <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
              {PROFILES.map(p => (
                <button key={p.value} className={`profile-option${profile === p.value ? " selected" : ""}`} onClick={() => handleProfileSelect(p.value)}>
                  {profile === p.value ? "✓ " : ""}{p.label}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* TUTORIAL */}
      {showTutorial && <TutorialModal onClose={closeTutorial} />}

      {/* LOGIN MODAL */}
      {showLoginModal && (
        <div className="modal-overlay" onClick={() => setShowLoginModal(false)}>
          <div className="modal" style={{ maxWidth: 340 }} onClick={e => e.stopPropagation()}>
            <button onClick={() => setShowLoginModal(false)} style={{ position: "absolute", top: 14, right: 14, background: "none", border: "none", color: "#a09888", fontSize: 18, cursor: "pointer" }}>✕</button>
            <h2 style={{ margin: "0 0 18px", fontFamily: "'Playfair Display', serif", fontSize: 19, color: "#f5efe0" }}>Accès administrateur</h2>
            <input type="password" placeholder="Mot de passe" value={passwordInput} onChange={e => setPasswordInput(e.target.value)} onKeyDown={e => e.key === "Enter" && handleLogin()}
              style={{ width: "100%", padding: "11px 14px", borderRadius: 10, border: "1px solid rgba(255,255,255,0.2)", background: "rgba(255,255,255,0.08)", color: "#f5efe0", fontSize: 14, fontFamily: "'Source Serif 4', serif", outline: "none", marginBottom: 8 }} />
            {loginError && <p style={{ color: "#ff8a80", fontSize: 12, margin: "0 0 8px" }}>{loginError}</p>}
            <button onClick={handleLogin} style={{ width: "100%", padding: "11px", borderRadius: 10, border: "none", background: "linear-gradient(135deg, #4fc3f7, #ce93d8)", color: "#fff", fontFamily: "'Source Serif 4', serif", fontSize: 14, fontWeight: "bold", cursor: "pointer" }}>Se connecter</button>
          </div>
        </div>
      )}
    </div>
  );
}
