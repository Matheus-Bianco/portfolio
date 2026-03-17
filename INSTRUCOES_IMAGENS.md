# Como adicionar fotos e arquivos ao portfólio

---

## Foto de perfil e CV

### Foto de perfil (página Sobre)
- Coloque sua foto em: `assets/images/profile.jpg` ou `assets/images/profile.jpeg`
- Formato recomendado: JPG/JPEG ou PNG
- **Redimensionamento automático:** execute `npm run resize-images` para ajustar todas as imagens ao padrão do site
- Se o arquivo não existir, será exibido um placeholder com ícone

### Currículo (CV)
- Coloque seu PDF em: `assets/files/cv.pdf`
- O botão "Baixar CV" na página Sobre aponta para esse arquivo
- Crie a pasta `assets/files/` se não existir

### Links LinkedIn e GitHub
- Já configurados em todas as páginas (navbar e rodapé)

---

## 1. Estrutura de pastas (screenshots dos projetos)

Crie a seguinte estrutura dentro da pasta `assets/images/projects/`:

```
assets/
└── images/
    └── projects/
        ├── project1/    ← Plataforma de Indicadores de Risco
        ├── project2/    ← Painel de Estudantes Imigrantes
        ├── project3/    ← Sistema GA/SE (SED Sergipe)
        ├── project4/    ← Sistema de Avaliações Diagnósticas
        ├── project5/    ← Painel de Recomposição
        ├── project6/    ← Monitoramento Temporal de Risco
        ├── project7/    ← Priorização SAEB
        └── project8/    ← Painel IDEB & Censo
```

**As pastas `project1` a `project8` já existem.** Basta colocar as imagens dentro delas.

---

## 2. Nomenclatura dos arquivos

Use os nomes **1.jpg**, **2.jpg** e **3.jpg** em cada pasta de projeto:

| Projeto | Arquivos esperados |
|---------|--------------------|
| **project1** | `1.jpg`, `2.jpg`, `3.jpg` |
| **project2** | `1.jpg`, `2.jpg`, `3.jpg` |
| **project3** | `1.jpg`, `2.jpg`, `3.jpg` |
| **project4** | `1.jpg`, `2.jpg`, `3.jpg`, `4.jpg` |
| **project5** | `1.jpg`, `2.jpg`, `3.jpg` |
| **project6** | `1.png`, `2.png`, `3.png` |
| **project7** | `1.jpg`, `2.jpg`, `3.jpg` |
| **project8** | `1.png`, `2.png`, `3.png`, `4.png`, `5.png` |

- **1.jpg** → Imagem em destaque (largura total, aparece primeiro) e **capa do card** na página de Projetos
- **2.jpg**, **3.jpg** (e **4.jpg** no project4) → Imagens menores ao lado

---

## 3. Formatos aceitos

- **Recomendado:** `.jpg` ou `.jpeg` (melhor para fotos e telas)
- **Alternativa:** `.png` (se precisar de transparência)

Se usar `.png`, altere a extensão no HTML. Exemplo:
```html
<img src="../assets/images/projects/project1/1.png" ...>
```

---

## 4. Tamanho e qualidade

Você **não precisa** redimensionar manualmente. Duas opções:

**Opção A – Script automático (recomendado):**
```bash
npm install
npm run resize-images
```
O script redimensiona: foto de perfil → 400×400 px; screenshots → 1200 px de largura.

**Opção B – CSS:** O layout já exibe qualquer tamanho de imagem corretamente (`object-fit: cover`).

---

## 5. Ordem das imagens

Coloque a melhor ou mais representativa do sistema como `1.jpg`.

---

## 6. Adicionar ou remover imagens

Para **adicionar** mais imagens em um projeto, copie este bloco no HTML da página do projeto (dentro de `<div class="project-screenshots">`):

```html
<figure class="project-screenshot">
  <img src="../assets/images/projects/project1/4.jpg" alt="Descrição da imagem">
  <figcaption>Legenda que aparece abaixo da foto</figcaption>
</figure>
```

Para **imagem em destaque** (largura total), use:
```html
<figure class="project-screenshot project-screenshot--featured">
  ...
</figure>
```

Para **remover** uma imagem, apague o bloco `<figure>...</figure>` correspondente do HTML.

---

## 7. Exemplo prático

Para o **Projeto 1** (Plataforma de Indicadores de Risco):

1. Tire 3 prints do seu sistema (dashboard, telas principais, etc.)
2. Salve como `1.jpg`, `2.jpg`, `3.jpg`
3. Coloque os arquivos em: `assets/images/projects/project1/`
4. Pronto. As imagens aparecerão automaticamente na página do projeto.

---

## 8. Se a imagem não aparecer

- Confira se o **caminho** está correto: `assets/images/projects/projectX/`
- Confira se o **nome do arquivo** está exato (incluindo maiúsculas/minúsculas)
- Em alguns projetos você pode ter menos de 3 imagens — nesse caso, remova os blocos `<figure>` das imagens que não existem no HTML da página.

---

## 9. Projeto sem screenshots ainda

Se um projeto ainda não tiver prints, você pode:
- **Opção A:** Deixar a seção como está — aparecerão ícones de imagem quebrada até você adicionar os arquivos
- **Opção B:** Remover toda a seção "Imagens do Sistema" daquele projeto (o bloco `<div class="project-section">` que contém `<h2>Imagens do Sistema</h2>` e o `<div class="project-screenshots">`)
