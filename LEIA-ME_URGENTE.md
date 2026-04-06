# 🚀 SOLUÇÃO RÁPIDA - Sistema não está funcionando

## O QUE FAZER AGORA:

### 1️⃣ Execute o Diagnóstico Automático

O sistema agora tem uma **função de diagnóstico** que vai mostrar exatamente qual é o problema!

**Como executar:**

1. Abra o **Apps Script Editor** do seu projeto
2. No topo da página, há um dropdown dizendo "Selecione uma função"
3. Clique e selecione: **`diagnosticoSistema`**
4. Clique no botão **▶️ Executar** (ao lado do dropdown)
5. Aguarde alguns segundos
6. Vá em **"Execuções"** no menu lateral esquerdo
7. Clique na execução mais recente
8. **LEIA O RELATÓRIO COMPLETO** que aparece no log

---

## 📊 O Diagnóstico Vai Verificar:

✅ Se seu email está sendo detectado  
✅ Se você está autorizado como coordenação  
✅ Se a pasta do Drive está acessível  
✅ Se a planilha foi criada ou se há erro  
✅ Se os cadastros estão funcionando  
✅ Qual é o seu perfil de acesso  

---

## 🔍 O Que Procurar no Relatório:

### ✅ Verde = OK
Tudo funcionando nessa área

### ⚠️  Amarelo = Aviso
Pode ser normal (ex: "nenhum professor cadastrado ainda")

### ❌ Vermelho = ERRO CRÍTICO
**Essa é a causa do problema!** Leia a mensagem e a solução sugerida.

---

## 🎯 Problemas Mais Comuns:

### Se aparecer: "Email não detectado" ou "Email inválido"
**Causa:** Você não está acessando via Web App implantado

**Solução:**
1. No Apps Script, vá em **Implantar** → **Gerenciar implantações**
2. Veja se tem uma implantação ativa
3. Se NÃO tiver: **Nova implantação** → **Web App**
   - Executar como: **Eu**
   - Quem tem acesso: **Qualquer pessoa**
4. Copie a **URL do Web App** e acesse ESSA URL no navegador
5. Autorize as permissões quando solicitado

---

### Se aparecer: "Não é coordenação"
**Causa:** Seu email não está na lista de emails autorizados

**Solução:**
1. Abra o arquivo `Codigo_v9.gs`
2. Procure por `EMAILS_COORDENACAO` (linhas 40-60)
3. Adicione seu email exatamente como aparece no relatório:
   ```javascript
   EMAILS_COORDENACAO: ["josival.lima@edu.mucuri.ba.gov.br"],
   ```
4. Salve (Ctrl+S)
5. Acesse o sistema novamente

---

### Se aparecer: "Erro ao acessar pasta"
**Causa:** ID da pasta está incorreto ou você não tem permissão

**Solução:**
1. Abra a pasta no Google Drive
2. Veja a URL: `https://drive.google.com/drive/folders/ESTE_É_O_ID`
3. Copie o ID (a parte depois de `/folders/`)
4. No `Codigo_v9.gs`, procure `PASTA_ID:`
5. Cole o ID correto:
   ```javascript
   PASTA_ID: "1oElOAG9HwoPTyEatRNskrgBHLPBXC3j1",
   ```
6. Salve e tente novamente

---

### Se aparecer: "Erro ao criar planilha"
**Causa:** Falta de permissões ou espaço no Drive

**Solução:**
1. Verifique se você tem espaço disponível no Google Drive
2. Certifique-se de que autorizou TODAS as permissões solicitadas
3. Se as permissões foram negadas anteriormente:
   - Vá em: https://myaccount.google.com/permissions
   - Remova o aplicativo "Sistema de Planejamento"
   - Execute o sistema novamente
   - Autorize TODAS as permissões quando solicitadas

---

## 🔄 Depois de Corrigir:

1. **Execute `diagnosticoSistema()` novamente**
2. Verifique se todos os testes estão com ✅
3. Se sim, acesse a URL do Web App
4. O sistema deve funcionar normalmente

---

## 📌 Informação Importante:

### Você DEVE acessar via Web App, NÃO pelo editor!

❌ **ERRADO:** `https://script.google.com/...` (editor do Apps Script)  
✅ **CERTO:** `https://script.google.com/macros/s/...` (URL do Web App)

A URL correta do Web App:
- Começa com `https://script.google.com/macros/s/`
- É gerada ao fazer deploy como "Web App"
- Fica em: **Implantar → Gerenciar implantações → URL**

---

## 🆘 Se Ainda Não Funcionar:

Copie **TODO O RELATÓRIO** da função `diagnosticoSistema()` e envie para análise.

O relatório mostra informações detalhadas sobre:
- Qual teste falhou
- Mensagem de erro completa
- Stack trace (se houver)
- Sugestões de solução

---

## ⚡ Resumo Rápido:

```
1. Execute diagnosticoSistema() no Apps Script
2. Veja qual teste tem ❌
3. Siga a solução específica para aquele teste
4. Execute diagnosticoSistema() novamente para confirmar
5. Acesse a URL do Web App (não o editor!)
```

---

**Boa sorte! 🍀**

Se seguir esses passos, o diagnóstico vai mostrar exatamente o que está errado e como corrigir!
