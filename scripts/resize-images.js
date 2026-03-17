/**
 * Script para redimensionar imagens do portfólio
 * 
 * Uso: npm run resize-images
 * 
 * Redimensiona automaticamente:
 * - Foto de perfil (assets/images/profile.*) → 400x400 px
 * - Screenshots dos projetos (assets/images/projects/projectX/*) → 1200px de largura
 */

const fs = require('fs');
const path = require('path');

async function resizeImages() {
  let sharp;
  try {
    sharp = require('sharp');
  } catch (e) {
    console.log('Instalando dependência "sharp"...');
    const { execSync } = require('child_process');
    execSync('npm install sharp', { stdio: 'inherit' });
    sharp = require('sharp');
  }

  const baseDir = path.join(__dirname, '..', 'assets', 'images');
  const ext = ['.jpg', '.jpeg', '.png', '.webp'];

  const processImage = async (inputPath, outputPath, options) => {
    const tempPath = outputPath + '.tmp';
    const opts = options.resizeOpts || {};
    const pipeline = options.height
      ? sharp(inputPath).resize(options.width, options.height, opts)
      : sharp(inputPath).resize(options.width, null, opts);
    const fileExt = path.extname(outputPath).toLowerCase();
    if (fileExt === '.png') {
      await pipeline.png({ quality: 90 }).toFile(tempPath);
    } else {
      await pipeline.jpeg({ quality: 85 }).toFile(tempPath);
    }
    fs.renameSync(tempPath, outputPath);
  };

  // Foto de perfil: 400x400 (mantém formato original)
  const profileFiles = ['profile.jpg', 'profile.jpeg', 'profile.png'];
  for (const file of profileFiles) {
    const filePath = path.join(baseDir, file);
    if (fs.existsSync(filePath)) {
      try {
        await processImage(filePath, filePath, {
          width: 400,
          height: 400,
          resizeOpts: { fit: 'cover', position: 'center' }
        });
        console.log('✓ Foto de perfil redimensionada:', file);
      } catch (err) {
        console.error('Erro em', file, err.message);
      }
    }
  }

  // Screenshots dos projetos: 1200px largura (mantém formato original)
  const projectsDir = path.join(baseDir, 'projects');
  if (fs.existsSync(projectsDir)) {
    const projectFolders = fs.readdirSync(projectsDir).filter(f => 
      fs.statSync(path.join(projectsDir, f)).isDirectory()
    );
    for (const project of projectFolders) {
      const projectPath = path.join(projectsDir, project);
      const files = fs.readdirSync(projectPath).filter(f => 
        ext.includes(path.extname(f).toLowerCase())
      );
      for (const file of files) {
        const filePath = path.join(projectPath, file);
        try {
          await processImage(filePath, filePath, {
            width: 1200,
            resizeOpts: { withoutEnlargement: true }
          });
          console.log('✓', project + '/' + file);
        } catch (err) {
          console.error('Erro em', project + '/' + file, err.message);
        }
      }
    }
  }

  console.log('\nConcluído!');
}

resizeImages().catch(err => {
  console.error(err);
  process.exit(1);
});
