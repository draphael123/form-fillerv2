// Run with: node generate-icons.js
// Requires: npm install canvas

const { createCanvas } = require('canvas');
const fs = require('fs');

[16, 48, 128].forEach(size => {
  const canvas = createCanvas(size, size);
  const ctx = canvas.getContext('2d');

  // Background
  const radius = size * 0.22;
  ctx.beginPath();
  ctx.moveTo(radius, 0);
  ctx.lineTo(size - radius, 0);
  ctx.quadraticCurveTo(size, 0, size, radius);
  ctx.lineTo(size, size - radius);
  ctx.quadraticCurveTo(size, size, size - radius, size);
  ctx.lineTo(radius, size);
  ctx.quadraticCurveTo(0, size, 0, size - radius);
  ctx.lineTo(0, radius);
  ctx.quadraticCurveTo(0, 0, radius, 0);
  ctx.closePath();

  const gradient = ctx.createLinearGradient(0, 0, size, size);
  gradient.addColorStop(0, '#4f7cff');
  gradient.addColorStop(1, '#7c4fff');
  ctx.fillStyle = gradient;
  ctx.fill();

  // Letter D
  ctx.fillStyle = '#ffffff';
  ctx.font = `bold ${size * 0.55}px Arial`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('D', size / 2, size / 2);

  fs.writeFileSync(`assets/icon${size}.png`, canvas.toBuffer('image/png'));
  console.log(`Generated icon${size}.png`);
});
