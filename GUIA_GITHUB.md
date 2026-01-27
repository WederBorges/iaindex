# 🚀 Guia para Publicar no GitHub

## Passo 1: Criar Repositório no GitHub

1. Acesse [GitHub](https://github.com) e faça login
2. Clique em **"New repository"** (ou **"+"** → **"New repository"**)
3. Configure:
   - **Repository name**: `iaindex` (ou outro nome de sua preferência)
   - **Description**: "Análise de termos em PDFs - IA vs Dados/BI"
   - **Visibility**: Público ou Privado (sua escolha)
   - **NÃO** marque "Initialize with README" (já temos um)
4. Clique em **"Create repository"**

## Passo 2: Adicionar Arquivos e Fazer Commit

Execute os seguintes comandos no terminal (na pasta do projeto):

```bash
# Adicionar todos os arquivos relevantes
git add .gitignore README.md LICENSE requirements.txt analisar_pdfs.py listar_empresas.py

# Fazer o commit inicial
git commit -m "Initial commit: Script de análise de termos IA vs Dados/BI em PDFs"
```

## Passo 3: Conectar ao Repositório Remoto

**Substitua `SEU-USUARIO` pelo seu nome de usuário do GitHub:**

```bash
# Adicionar o repositório remoto
git remote add origin https://github.com/SEU-USUARIO/iaindex.git

# Verificar se foi adicionado corretamente
git remote -v
```

## Passo 4: Fazer Push para o GitHub

```bash
# Enviar para o GitHub (primeira vez)
git branch -M main
git push -u origin main
```

Se você já configurou autenticação no GitHub (SSH ou token), o push funcionará. Caso contrário, você precisará:

### Opção A: Usar Personal Access Token
1. Vá em GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Crie um novo token com permissão `repo`
3. Use o token como senha quando solicitado

### Opção B: Usar SSH
```bash
# Alterar para SSH (se preferir)
git remote set-url origin git@github.com:SEU-USUARIO/iaindex.git
```

## ✅ Verificação

Após o push, acesse seu repositório no GitHub e verifique se todos os arquivos foram enviados corretamente.

## 📝 Próximos Commits

Para futuras atualizações:

```bash
# Adicionar mudanças
git add .

# Fazer commit
git commit -m "Descrição das mudanças"

# Enviar para o GitHub
git push
```

## 🔧 Comandos Úteis

```bash
# Ver status dos arquivos
git status

# Ver histórico de commits
git log

# Ver diferenças não commitadas
git diff
```
