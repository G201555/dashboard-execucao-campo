// ============================================================
// Publica DADOS/DADOS_VENDAS (extraidos do dashboard.html local)
// no Firestore do projeto seu-projeto-firebase, autenticado
// como usuario autorizado (via REST, sem depender de service
// account / gcloud). Roda so quando voce quer atualizar a versao
// publicada no Firebase.
//
// Uso:
//   node publicar_firestore.js <email> <senha>
// ============================================================
const fs = require('fs');
const path = require('path');

const API_KEY = 'SUA_API_KEY';
const PROJECT_ID = 'seu-projeto-firebase';
const CHUNK_SIZE = 150; // ~150 lojas por doc, bem abaixo do limite de 1MiB do Firestore

const email = process.argv[2];
const senha = process.argv[3];
if (!email || !senha) {
  console.error('Uso: node publicar_firestore.js <email> <senha>');
  process.exit(1);
}

function extrairBloco(html, marcadorInicio, marcadorFim, prefixo) {
  const re = new RegExp(
    marcadorInicio.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '([\\s\\S]*?)' + marcadorFim.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  );
  const m = html.match(re);
  if (!m) throw new Error(`Bloco ${marcadorInicio} não encontrado`);
  const bloco = m[1];
  const jm = bloco.match(new RegExp(prefixo + '(\\{.*\\});'));
  if (!jm) throw new Error(`JSON de ${prefixo} não encontrado`);
  return JSON.parse(jm[1]);
}

async function main() {
  const htmlPath = path.join(__dirname, '..', 'dashboard.html');
  const html = fs.readFileSync(htmlPath, 'utf8');
  const DADOS = extrairBloco(html, '/*__DADOS_START__*/', '/*__DADOS_END__*/', 'const DADOS=');
  const DADOS_VENDAS = extrairBloco(html, '/*__VENDAS_START__*/', '/*__VENDAS_END__*/', 'const DADOS_VENDAS=');
  console.log(`Lidos ${DADOS.lojas.length} lojas de campo e ${DADOS_VENDAS.lojas.length} lojas de vendas.`);

  // ---------- login ----------
  const signInResp = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${API_KEY}`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password: senha, returnSecureToken: true }),
  });
  const signInData = await signInResp.json();
  if (!signInData.idToken) throw new Error('Falha no login: ' + JSON.stringify(signInData));
  const idToken = signInData.idToken;
  console.log('Login OK como', signInData.email);

  // ---------- helper: escreve 1 doc com campo json (string) ----------
  async function setDoc(docId, jsonString) {
    const url = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/nintendo/${docId}`;
    const resp = await fetch(url, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${idToken}` },
      body: JSON.stringify({ fields: { json: { stringValue: jsonString } } }),
    });
    const data = await resp.json();
    if (data.error) throw new Error(`Falha ao gravar ${docId}: ${JSON.stringify(data.error)}`);
    return data;
  }

  // ---------- monta chunks de lojas, um jogo de chunks por mes ----------
  const porMes = DADOS.porMes || { atual: { lojas: DADOS.lojas } };
  const meses = DADOS.meses && DADOS.meses.length ? DADOS.meses : [{ chave: 'atual', rotulo: 'Atual' }];
  const numChunksPorMes = {};
  const chunksPorMes = {};
  for (const m of meses) {
    const lojasDoMes = (porMes[m.chave] || {}).lojas || [];
    const chunks = [];
    for (let i = 0; i < lojasDoMes.length; i += CHUNK_SIZE) chunks.push(lojasDoMes.slice(i, i + CHUNK_SIZE));
    chunksPorMes[m.chave] = chunks;
    numChunksPorMes[m.chave] = chunks.length;
  }

  const meta = {
    geradoEm: DADOS.geradoEm,
    janelaMixDias: DADOS.janelaMixDias,
    tendencia: DADOS.tendencia,
    skus: DADOS.skus,
    campanhas: DADOS.campanhas || [],
    campanhasDetalhe: DADOS.campanhasDetalhe || [],
    ocorrencias: DADOS.ocorrencias || [],
    meses,
    mesAtual: DADOS.mesAtual || 'atual',
    numChunksPorMes,
  };

  const totalChunks = Object.values(numChunksPorMes).reduce((a, b) => a + b, 0);
  console.log(`Gravando meta + ${totalChunks} chunks (${meses.length} meses) + vendas...`);
  await setDoc('meta', JSON.stringify(meta));
  let feitos = 0;
  for (const m of meses) {
    const chunks = chunksPorMes[m.chave];
    for (let i = 0; i < chunks.length; i++) {
      await setDoc(`chunk_${m.chave}_${i}`, JSON.stringify(chunks[i]));
      feitos++;
      process.stdout.write(`  chunk ${feitos}/${totalChunks} (${m.rotulo})\r`);
    }
  }
  console.log('\n  vendas...');
  // fatos (loja x produto x semana) crescem toda semana: vao em docs proprios pra nao estourar 1 MB
  const fatos = DADOS_VENDAS.fatos || [];
  const FATOS_POR_DOC = 15000;
  const numChunksFatos = Math.ceil(fatos.length / FATOS_POR_DOC);
  await setDoc('vendas', JSON.stringify({ ...DADOS_VENDAS, fatos: [], numChunksFatos }));
  for (let i = 0; i < numChunksFatos; i++) {
    await setDoc(`vendas_fatos_${i}`, JSON.stringify(fatos.slice(i * FATOS_POR_DOC, (i + 1) * FATOS_POR_DOC)));
  }
  console.log(`  vendas: ${fatos.length} linhas em ${numChunksFatos} doc(s)`);
  // doc proprio: ~1.400 lojas da planilha do roteiro nao cabem com folga junto do meta
  console.log('  base de lojas...');
  await setDoc('baseLojas', JSON.stringify(DADOS.baseLojas || null));

  console.log('OK - dados publicados no Firestore.');
}

main().catch(e => { console.error(e); process.exit(1); });
