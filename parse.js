const fs = require('fs');
const xml2js = require('xml2js');

const showImage = (image, subFolder, title) => {
  if (image === undefined) return '';
  return `\n![Screenshot of ${title}](${subFolder}/screenshot.${image})\n\n`;
}
const stats = {};

function createMarkdown(scripts) {
  scripts = scripts.map(s => s['$']);

  console.log(scripts.length); // 220
  scripts.sort((a, b) => {
    const aa = new Date(a.date);
    const bb = new Date(b.date);
    if (aa < bb) return -1;
    if (bb > aa) return 1;
    return 0;
  });
  console.log(scripts[0].date);
  console.log(scripts[scripts.length - 1].date);

  for (const script of scripts) {
    const {
      title,
      description,
      date,
      image,
      code,
      language,
      folder
    } = script;
    if (language in stats) {
      stats[language]++;
    } else {
      stats[language] = 1;
    }
    const [topDir, subFolder] = folder.split('/');

    const markdown = `## [${title}](./${subFolder})

*${date}*

${description}
${showImage(image, subFolder, title)}

`;


    fs.appendFileSync(`${topDir}/README.md`, markdown, 'utf8');
  }
  console.log(stats);
}

const names = ['SQL', 'Visual Basic', 'VB.Net', 'JavaScript', 'C#', 'C/C++', 'ASP.Net', 'Classic ASP / vbScript'];
['SQL', 'VB', 'VBNet', 'JavaScript', 'CSharp', 'C', 'ASPNet', 'ASP']
  .forEach((topDir, i) => {

    let header = `# Lewie's Code Library PSC

## ${names[i]}

Open source projects that I had published to Planet Source Code.

`;

    fs.writeFileSync(`${topDir}/README.md`, header, 'utf8');
  });


const xml = fs.readFileSync('scripts.xml', 'utf8');

xml2js.parseString(xml, (err, result) => {
  if (err) {
    console.error('Error parsing XML:', err);
    return;
  }
  createMarkdown(result.scripts.script);
});
