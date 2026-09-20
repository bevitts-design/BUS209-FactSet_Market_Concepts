// Rebuild the infographic with genuine screenshot pixels, never generated UI.
// Run with the bundled Node runtime; set NODE_PATH to its node_modules directory.
const sharp = require('sharp');
const path = require('path');
const root = path.resolve(__dirname, '..');
const asset = name => path.join(root, 'assets', name);
async function build() {
  const base = asset('factset-10-stock-portfolio-base.png');
  const capture = asset('factset-creating-portfolio-help.png');
  const top = await sharp(base).extract({left:0, top:0, width:1024, height:208}).png().toBuffer();
  const body = await sharp(base).extract({left:0, top:208, width:1024, height:1328}).png().toBuffer();
  // Coordinates refer to the unaltered 1363 x 768 OA screenshot.
  const icon = await sharp(capture).extract({left:493, top:228, width:28, height:29})
    .resize(112,116,{kernel:'nearest'}).png().toBuffer();
  const dropdown = await sharp(capture).extract({left:465, top:331, width:72, height:46})
    .resize(288,184,{kernel:'nearest'}).png().toBuffer();
  await sharp(icon).toFile(asset('factset-create-edit-watchlist-icon.png'));
  await sharp(dropdown).toFile(asset('factset-ticker-shares-dropdown.png'));
  const strip = Buffer.from(`<svg width="1024" height="320" xmlns="http://www.w3.org/2000/svg">
    <rect width="1024" height="320" fill="white"/>
    <rect x="20" y="8" width="984" height="298" rx="16" fill="#edf5fb" stroke="#b9d9ef" stroke-width="2"/>
    <g font-family="Arial, sans-serif" fill="#0a2540">
      <text x="42" y="45" font-size="25" font-weight="bold">Recognize these controls</text>
      <text x="70" y="85" font-size="23" font-weight="bold">1  Create/Edit Watchlist</text>
      <text x="581" y="85" font-size="23" font-weight="bold">2  Ticker + Shares</text>
      <text x="70" y="252" font-size="19">Look for this toolbar icon.</text>
      <text x="42" y="290" font-size="17" fill="#586b7a">Actual crops from FactSet Online Assistant • enlarged for visibility</text>
    </g>
  </svg>`);
  await sharp({create:{width:1024,height:1856,channels:4,background:'white'}})
    .composite([{input:top,left:0,top:0},{input:strip,left:0,top:208},
      {input:icon,left:150,top:312},{input:dropdown,left:579,top:303},
      {input:body,left:0,top:528}])
    .png().toFile(asset('factset-10-stock-portfolio.png'));
}
build().catch(error => { console.error(error); process.exitCode=1; });
