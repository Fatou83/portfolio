const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
  ExternalHyperlink, TabStopType, TabStopPosition, BorderStyle, HeadingLevel
} = require("docx");

// ---- Helpers ----
function sectionTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
    children: [new TextRun({ text, bold: true, color: "2E75B6" })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 40 },
    children: [new TextRun({ text })],
  });
}

function jobHeader(title, dates) {
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: { before: 160, after: 0 },
    children: [
      new TextRun({ text: title, bold: true }),
      new TextRun({ text: `\t${dates}`, bold: true }),
    ],
  });
}

function companyLine(company, location) {
  return new Paragraph({
    spacing: { after: 60 },
    children: [
      new TextRun({ text: company, italics: true }),
      new TextRun({ text: location ? `, ${location}` : "", italics: true }),
    ],
  });
}

function bodyParagraph(text) {
  return new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text })] });
}

// ---- Content ----
const children = [];

// Header / name
children.push(new Paragraph({
  spacing: { after: 20 },
  children: [new TextRun({ text: "FATOU MANE NDIAYE", bold: true, size: 36 })],
}));

children.push(new Paragraph({
  spacing: { after: 100 },
  children: [new TextRun({ text: "Consultante informatique et Développeuse", size: 22, color: "2E75B6", bold: true })],
}));

children.push(new Paragraph({
  spacing: { after: 60 },
  children: [
    new TextRun({ text: "Sacrée Cœur, Dakar  |  +221-77-748-98-85  |  fatoumanendiaye83@gmail.com" }),
  ],
}));

children.push(new Paragraph({
  spacing: { after: 200 },
  children: [
    new ExternalHyperlink({ link: "https://www.linkedin.com/in/fatou-man%C3%A9-ndiaye-2a87b4167/", children: [new TextRun({ text: "LinkedIn", style: "Hyperlink" })] }),
    new TextRun({ text: "  |  " }),
    new ExternalHyperlink({ link: "https://fatou83.github.io/portfolio/", children: [new TextRun({ text: "https://fatou83.github.io/portfolio/", style: "Hyperlink" })] }),
    new TextRun({ text: "  |  Permis de conduire : B" }),
  ],
}));

// PROFIL
children.push(sectionTitle("PROFIL"));
children.push(bodyParagraph(
  "Développeuse web et mobile passionnée, spécialisée dans la création de solutions modernes et performantes. Forte d'une expertise en développement front-end et back-end, ainsi qu'en conception d'applications mobiles hybrides. Dotée d'une excellente maîtrise des technologies telles que Keycloak, JavaScript, Node.js, Angular, Java, Flutter et des bases de données comme MySQL et PostgreSQL. Capable de transformer des idées en produits numériques innovants, avec une approche centrée sur l'utilisateur et une grande attention aux détails. Dynamique, adaptable et orientée résultats, je suis motivée par les défis techniques et les projets qui ont un impact concret."
));

// EXPERIENCE
children.push(sectionTitle("EXPÉRIENCE PROFESSIONNELLE"));

// Job 1
children.push(jobHeader("Consultante informatique et Développeuse", "06/2025 – Présent"));
children.push(companyLine("CINERI", "Dakar, Sénégal"));
[
  "Mise en place d'une authentification centralisée.",
  "Création et installation d'un plugin d'envoi de code OTP pour l'authentification multi-facteur (MFA).",
  "Développement d'un thème personnalisé pour Keycloak.",
  "Conception d'une application de gestion des inscriptions des étudiants.",
  "Élaboration d'une application de gestion des orientations des étudiants.",
].forEach(t => children.push(bullet(t)));

// Job 2
children.push(jobHeader("Consultante informatique et Développeuse", "11/2024 – 02/2025"));
children.push(companyLine("TechCo", "Dakar, Sénégal"));
[
  "Création d'une application de gestion des bulletins de salaire.",
  "Intégration de l'API POSTMARK pour l'envoi d'e-mails.",
  "Intégration des APIs SMS et WhatsApp pour notifier les utilisateurs après l'envoi des bulletins.",
].forEach(t => children.push(bullet(t)));

// Job 3
children.push(jobHeader("Consultante informatique et Développeuse", "06/2024 – 11/2024"));
children.push(companyLine("Consi", "Remote, Sénégal"));
[
  "Développement d'une application mobile avec Flutter pour les patients.",
  "Intégration des APIs backend sur l'application mobile avec Flutter.",
  "Refonte du site web sur WordPress.",
  "Mise en œuvre des APIs SMS et WhatsApp pour notifier les utilisateurs avec Node.js.",
  "Génération d'un code OTP et vérification du code reçu.",
  "Conception d'une application de gestion d'annonces avec Flutter.",
  "Intégration de l'API PayPal sur le backend (Node.js) et le frontend (Flutter).",
].forEach(t => children.push(bullet(t)));

// Job 4
children.push(jobHeader("Développeuse Web et Mobile", "02/2022 – 06/2024"));
children.push(companyLine("Saou Transport", "Dakar"));
[
  "Développement d'une application mobile avec Ionic.",
  "Compilation de l'application sur Xcode pour obtenir la version iOS.",
  "Installation des plugins avec Capacitor.",
  "Refonte du design de l'application mobile avec Ionic.",
  "Ajout du module de gestion des colis dans la partie web à l'aide de Spring et Angular.",
  "Déploiement de l'application mobile sur l'App Store.",
  "Déploiement de l'application web à l'aide de Docker.",
  "Analyse des cahiers des charges pour comprendre les besoins des clients et anticiper les contraintes techniques (type de plateforme, taille de l'écran, degré d'interactivité, langage de programmation...).",
  "Identification de la solution technique la plus adaptée au projet.",
  "Prise en charge de la réalisation technique de l'application, en étroite collaboration avec le chef de projet et le responsable des opérations.",
  "Réalisation des tests et résolution des éventuels dysfonctionnements.",
  "Assurer la maintenance de l'application.",
  "Optimisation de l'application et intégration de nouvelles fonctionnalités.",
  "Conception de l'application mobile avec Flutter.",
  "Assurer le déploiement de l'application (installation, assistance, formation, évaluation).",
  "Maintenance de l'application (diagnostic des défauts, correction) et son évolution.",
].forEach(t => children.push(bullet(t)));

// Job 5
children.push(new Paragraph({
  spacing: { before: 160, after: 0 },
  children: [new TextRun({ text: "Développeuse web et mobile d'une application de transport de type IoT", bold: true })],
}));
children.push(new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
  spacing: { after: 60 },
  children: [
    new TextRun({ text: "Sn software, Dakar, Sénégal", italics: true }),
    new TextRun({ text: "\t06/2021 – 05/2022", bold: true }),
  ],
}));
[
  "Création d'interface utilisateur et intégration d'API à l'aide d'Angular et Ionic.",
  "Choix de l'architecture microservices.",
  "Développement des microservices et du gateway.",
  "Conception de l'application mobile avec Flutter.",
  "Ajout du module de gestion des colis dans la partie web à l'aide de Spring et Angular.",
  "Intégration de l'API d'une autre société de transport pour permettre aux utilisateurs d'acheter directement leurs billets via notre application, à l'aide de Java Spring Boot et Angular.",
  "Création du tableau de bord avec les statistiques (nombre de tickets vendus par mois, recette par jour).",
  "Localisation des bus et leur statut (en cours, en panne...) sur une carte.",
  "Développement d'une application pour une compagnie aérienne avec Odoo.",
  "Hébergement du site web sur un serveur.",
  "Gestion des tickets.",
  "Gestion des erreurs sur le site web et maintenance du serveur.",
].forEach(t => children.push(bullet(t)));

// FORMATION
children.push(sectionTitle("FORMATION"));

function eduEntry(school, degree, city, dates) {
  children.push(jobHeader(school, dates));
  children.push(bodyParagraph(`${degree} — ${city}`));
}

eduEntry("INSTITUT SUP'INFO", "Master, Spécialité programmation et développement web et mobile", "Dakar", "2019 – 2021");
eduEntry("INSTITUT SUP'INFO", "Licence 3, Spécialité programmation et développement web et mobile", "Dakar", "2017 – 2019");
eduEntry("INSTITUT SUPÉRIEUR INFORMATIQUE (ISI)", "Licence 3, Spécialité réseaux informatiques", "Dakar", "2015 – 2018");
eduEntry("UCAD", "Licence 1, Science physique", "Dakar", "2013 – 2014");
eduEntry("LCND", "Bac S2", "Fatick", "2012 – 2013");

// CONNAISSANCES
children.push(sectionTitle("CONNAISSANCES"));
children.push(bodyParagraph(
  "Java, Angular, Spring, Hibernate, JPA, Ionic, Flutter, Keycloak, JavaScript, Node.js, MySQL, PostgreSQL, Maven, JHipster, Linux, Windows, macOS, API REST, GitHub, Docker, App Store, Play Store, CRM, CMS, WordPress, Odoo, Drupal 8, Drupal 9, Java Web, HTML, XML, API Postmark, DocuSign, Conception et administration de site, Conception de base de données, SQLite, Oracle."
));
children.push(bodyParagraph("Qualités : orientée résultats, adaptable, dynamique."));

// LANGUES
children.push(sectionTitle("LANGUES"));
[
  "Français — courant",
  "Anglais — intermédiaire (B1)",
  "Wolof — courant",
].forEach(t => children.push(bullet(t)));

// CERTIFICATS
children.push(sectionTitle("CERTIFICATS"));
children.push(bullet("Permis de conduire : B"));

// ---- Document ----
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("CV_Fatou_Mane_Ndiaye.docx", buffer);
  console.log("done");
});
