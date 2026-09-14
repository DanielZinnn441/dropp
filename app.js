/* ==========================================================================
   ArbitragePulse — Planilha de Viabilidade e Margem de Produtos (Dropship)
   ========================================================================== */

const STORAGE_KEY = 'planilha_produtos_dropship_v4';

// Dados reais extraídos das pastas PROFUTO 01 a PRODUTO 05
const INITIAL_PRODUCTS = [
  {
    id: 'prod-001',
    date: '10/09/2026',
    name: 'Kit 3 Potes Organizadores para Geladeira Premium com Cestos Escorredores',
    category: 'Cozinha / Organização',
    amzPrice: 43.90,
    amzLink: 'https://www.amazon.com.br/Organizadores-Geladeira-Escorredores-Herm%C3%A9ticas-Conserva/dp/B0H9P9QR57',
    amzReviews: 1,
    amzRating: 5.0,
    amzComp: 1,
    
    shpStore: 'Shopee (Cruzeiro, SP)',
    shpLink: 'https://shopee.com.br/Kit-3-potes-Organizadores-Geladeira-Organizador-Alimentos-Vegetais-Cozinha-Moderna-i.629805123.58205197736',
    sourcePrice: 32.98,
    freightEst: 12.04,
    amazonTaxPct: 0.15,
    otherCosts: 0.00,
    
    isFavorite: true,
    priority: 'Alta',
    notes: 'Nota 5,0 na Amazon. Com cupom frete grátis Shopee o lucro fica em +R$ 4,33 (9,86%).'
  },
  {
    id: 'prod-002',
    date: '10/09/2026',
    name: 'Pré Treino Insane Clown 350g Demons Lab - Blue Crystal',
    category: 'Suplementos / Saúde',
    amzPrice: 116.90,
    amzLink: 'https://www.amazon.com.br/Treino-Insane-Clown-350g-Demons/dp/B0GYQM5WFT',
    amzReviews: 'N/I',
    amzRating: 'N/I',
    amzComp: 'N/I',
    
    shpStore: 'Shopee (Cruzeiro, SP)',
    shpLink: 'https://shopee.com.br/INSANE-CLOWN-350G-BLUE-CRYSTAL-i.384424516.22391686410',
    sourcePrice: 99.75,
    freightEst: 10.83,
    amazonTaxPct: 0.15,
    otherCosts: 0.00,
    
    isFavorite: false,
    priority: 'Média',
    notes: 'Recomenda-se reajustar preço na Amazon para R$ 139,90 para garantir margem positiva.'
  },
  {
    id: 'prod-003',
    date: '10/09/2026',
    name: 'Kit 6 Pares Meias Cano Invisível Lupo Sport Sapatilha',
    category: 'Moda / Roupas',
    amzPrice: 74.90,
    amzLink: 'https://www.amazon.com.br/Lupo-esportivas-algod%C3%A3o-respir%C3%A1vel-pares/dp/B00PROD03',
    amzReviews: 1449,
    amzRating: 4.8,
    amzComp: 3,
    
    shpStore: 'Shopee Lojas Oficiais',
    shpLink: 'https://shopee.com.br/Kit-De-6-Pares-Meias-Cano-Invis%C3%ADvel-Lupo-Sport-Algod%C3%A3o-Soquete-Sapatilha-Unissex',
    sourcePrice: 49.23,
    freightEst: 9.62,
    amazonTaxPct: 0.15,
    otherCosts: 0.00,
    
    isFavorite: true,
    priority: 'Alta',
    notes: 'Alta demanda na Amazon (4,8★ com 1,4k+ avaliações). Lucro sobe para R$ 14,44 (19,27%) com frete grátis.'
  },
  {
    id: 'prod-004',
    date: '10/09/2026',
    name: 'Carregador Turbo Duo 168W USB-C e USB-A com Cabo USB-C',
    category: 'Eletrônicos / Acessórios',
    amzPrice: 79.90,
    amzLink: 'https://www.amazon.com.br/dp/B0H5XWN3WD',
    amzReviews: 'N/I',
    amzRating: 'N/I',
    amzComp: 'N/I',
    
    shpStore: 'Shopee (Cruzeiro, SP)',
    shpLink: 'https://shopee.com.br/product/1572491828/22094298275',
    sourcePrice: 25.88,
    freightEst: 9.62,
    amazonTaxPct: 0.15,
    otherCosts: 0.00,
    
    isFavorite: true,
    priority: 'Alta',
    notes: 'Produto Destaque! Lucro de R$ 32,42 a R$ 42,04 por unidade (Margem > 40% e ROI até 162%).'
  },
  {
    id: 'prod-005',
    date: '10/09/2026',
    name: 'Jogo de Talheres Faqueiro Aço Inox 24 Peças',
    category: 'Cozinha / Organização',
    amzPrice: 52.69,
    amzLink: 'https://www.amazon.com.br/Jogo-Talheres-Faqueiro-Inox-Pecas/dp/B00PROD05',
    amzReviews: 4,
    amzRating: 2.6,
    amzComp: 1,
    
    shpStore: 'Shopee (Cruzeiro, SP)',
    shpLink: 'https://shopee.com.br/Jogo-de-Talheres-24-Pe%C3%A7as-Faqueiro-A%C3%A7o-Inox-Faca-Garfo-Colher-de-Mesa-e-Colheres-de-Sobremesa',
    sourcePrice: 25.40,
    freightEst: 9.62,
    amazonTaxPct: 0.15,
    otherCosts: 0.00,
    
    isFavorite: false,
    priority: 'Média',
    notes: 'Boa margem de lucro (18% a 36%), porém requer atenção pela nota 2,6 na Amazon.'
  }
];

// ESTADO DA APLICAÇÃO
let products = [];
let currentFilter = {
  search: '',
  category: 'ALL',
  minMargin: 0,
  favoritesOnly: false
};
let isFreeFreightCoupon = false;
let activeTab = 'ALL_PRODUCTS';

// INICIALIZAÇÃO
document.addEventListener('DOMContentLoaded', () => {
  loadData();
  setupEventListeners();
  renderApp();
});

function loadData() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try {
      products = JSON.parse(saved);
    } catch (e) {
      console.error('Erro ao carregar dados do localStorage', e);
      products = INITIAL_PRODUCTS;
    }
  } else {
    products = INITIAL_PRODUCTS;
    saveData();
  }
}

function saveData() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(products));
}

// CALCULADORA DE FINANÇAS E MARGEM
function calculateFinancials(item) {
  const freightActual = isFreeFreightCoupon ? 0 : (item.freightEst || 0);
  const costDelivered = (item.sourcePrice || 0) + freightActual;
  const sellPrice = item.amzPrice || 0;
  const taxAmazonR$ = sellPrice * (item.amazonTaxPct || 0.15);
  const totalCostForSale = costDelivered + taxAmazonR$ + (item.otherCosts || 0);
  
  const netProfit = sellPrice - totalCostForSale;
  const profitMarginPct = sellPrice > 0 ? (netProfit / sellPrice) : 0;
  const roiPct = costDelivered > 0 ? (netProfit / costDelivered) : 0;

  return {
    freightActual,
    costDelivered,
    sellPrice,
    taxAmazonR$,
    totalCostForSale,
    netProfit,
    profitMarginPct,
    roiPct
  };
}

// RENDERIZAÇÃO DE TELA E KPIS
function renderApp() {
  renderKPIs();
  renderTable();
  updateTabCounts();
}

function renderKPIs() {
  const totalProducts = products.length;
  let highestProfit = -Infinity;
  let highestRoi = -Infinity;
  let marginSum = 0;
  let validCount = 0;

  products.forEach(p => {
    if (p.amzPrice) {
      const fin = calculateFinancials(p);
      if (fin.netProfit > highestProfit) highestProfit = fin.netProfit;
      if (fin.roiPct > highestRoi) highestRoi = fin.roiPct;
      marginSum += fin.profitMarginPct;
      validCount++;
    }
  });

  const avgMargin = validCount > 0 ? (marginSum / validCount) : 0;

  const elTotal = document.getElementById('kpi-total-prod');
  const elProfit = document.getElementById('kpi-best-profit');
  const elRoi = document.getElementById('kpi-best-roi');
  const elAvg = document.getElementById('kpi-avg-margin');

  if (elTotal) elTotal.innerText = `${totalProducts} Produtos`;
  if (elProfit) elProfit.innerText = highestProfit > -Infinity ? formatBRL(highestProfit) : 'R$ 0,00';
  if (elRoi) elRoi.innerText = highestRoi > -Infinity ? formatPct(highestRoi) : '0,00%';
  if (elAvg) elAvg.innerText = formatPct(avgMargin);
}

function renderTable() {
  const tbody = document.getElementById('tableBody');
  if (!tbody) return;

  tbody.innerHTML = '';

  const filtered = products.filter(item => {
    // Filtro de Aba
    if (activeTab === 'HIGH_MARGIN') {
      const fin = calculateFinancials(item);
      if (fin.profitMarginPct < 0.15) return false;
    } else if (activeTab === 'FAVORITES') {
      if (!item.isFavorite) return false;
    }

    // Filtro de Busca
    if (currentFilter.search) {
      const q = currentFilter.search.toLowerCase();
      const matchName = (item.name || '').toLowerCase().includes(q);
      const matchCat = (item.category || '').toLowerCase().includes(q);
      if (!matchName && !matchCat) return false;
    }

    // Categoria
    if (currentFilter.category !== 'ALL' && item.category !== currentFilter.category) return false;

    // Margem Mínima
    if (currentFilter.minMargin > 0) {
      const fin = calculateFinancials(item);
      if (fin.profitMarginPct * 100 < currentFilter.minMargin) return false;
    }

    // Apenas Favoritos
    if (currentFilter.favoritesOnly && !item.isFavorite) return false;

    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="20" style="text-align: center; padding: 40px; color: var(--text-muted);">
          Nenhum produto encontrado com os filtros selecionados. Clique em "+ Adicionar Produto" para inserir novos itens!
        </td>
      </tr>
    `;
    return;
  }

  filtered.forEach(item => {
    const fin = calculateFinancials(item);
    const tr = document.createElement('tr');

    tr.innerHTML = `
      <td>
        <button onclick="toggleFavorite('${item.id}')" class="star-btn ${item.isFavorite ? 'active' : ''}">
          ${item.isFavorite ? '★' : '☆'}
        </button>
      </td>
      <td><small style="color:var(--text-muted);">${item.date || '10/09/2026'}</small></td>
      <td><strong>${escapeHtml(item.name)}</strong><br><small style="color:var(--text-muted);">${item.category}</small></td>
      <td>${item.amzLink ? `<a href="${item.amzLink}" target="_blank" style="color:var(--primary-blue); text-decoration:none;">Ver Anúncio ↗</a>` : '-'}</td>
      <td><strong>${formatBRL(item.amzPrice)}</strong></td>
      
      <td>${item.amzReviews || 'N/I'}</td>
      <td>${typeof item.amzRating === 'number' ? item.amzRating.toFixed(1) + ' ★' : (item.amzRating || 'N/I')}</td>
      <td>${item.amzComp || 'N/I'}</td>
      
      <td>${escapeHtml(item.shpStore || 'Shopee')}</td>
      <td>${item.shpLink ? `<a href="${item.shpLink}" target="_blank" style="color:var(--primary-blue); text-decoration:none;">Ver Loja ↗</a>` : '-'}</td>
      <td>${formatBRL(item.sourcePrice)}</td>
      <td>${formatBRL(fin.freightActual)} ${isFreeFreightCoupon ? '<small style="color:var(--accent-green)">(Cupom)</small>' : ''}</td>
      <td><strong>${formatBRL(fin.costDelivered)}</strong></td>
      
      <td>${formatBRL(fin.taxAmazonR$)} <small>(15%)</small></td>
      <td><strong>${formatBRL(fin.totalCostForSale)}</strong></td>
      
      <td class="${fin.netProfit >= 0 ? 'val-positive' : 'val-negative'}">${formatBRL(fin.netProfit)}</td>
      <td class="${fin.profitMarginPct >= 0 ? 'val-positive' : 'val-negative'}">${formatPct(fin.profitMarginPct)}</td>
      <td class="${fin.roiPct >= 0 ? 'val-positive' : 'val-negative'}">${formatPct(fin.roiPct)}</td>
      
      <td style="white-space: nowrap;">
        <button onclick="openEditModal('${item.id}')" class="btn btn-outline btn-sm" title="Editar Produto" style="margin-right:4px; color:var(--primary-blue); border-color: rgba(59,130,246,0.4);">✏️ Editar</button>
        <button onclick="deleteProduct('${item.id}')" class="btn btn-outline btn-sm" title="Excluir" style="color:var(--accent-red);">✕</button>
      </td>
    `;

    tbody.appendChild(tr);
  });
}

function updateTabCounts() {
  const favCount = products.filter(p => p.isFavorite).length;

  const elAll = document.getElementById('badge-count-all');
  const elFav = document.getElementById('badge-count-fav');

  if (elAll) elAll.innerText = products.length;
  if (elFav) elFav.innerText = favCount;
}

// LISTENERS DE EVENTOS
function setupEventListeners() {
  const searchEl = document.getElementById('searchInput');
  if (searchEl) searchEl.addEventListener('input', (e) => { currentFilter.search = e.target.value; renderTable(); });

  const catEl = document.getElementById('catFilter');
  if (catEl) catEl.addEventListener('change', (e) => { currentFilter.category = e.target.value; renderTable(); });

  const marginEl = document.getElementById('minMarginFilter');
  if (marginEl) marginEl.addEventListener('change', (e) => { currentFilter.minMargin = parseFloat(e.target.value) || 0; renderTable(); });

  const favEl = document.getElementById('favToggle');
  if (favEl) favEl.addEventListener('change', (e) => { currentFilter.favoritesOnly = e.target.checked; renderTable(); });

  const freightEl = document.getElementById('freightToggle');
  if (freightEl) freightEl.addEventListener('change', (e) => { isFreeFreightCoupon = e.target.checked; renderApp(); });

  const formAdd = document.getElementById('addProductForm');
  if (formAdd) formAdd.addEventListener('submit', handleAddProduct);

  const formEdit = document.getElementById('editProductForm');
  if (formEdit) formEdit.addEventListener('submit', handleEditProduct);
}

function switchTab(tabName) {
  activeTab = tabName;
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  const btnActive = document.getElementById(`tab-${tabName}`);
  if (btnActive) btnActive.classList.add('active');
  renderTable();
}

function toggleFavorite(id) {
  const prod = products.find(p => p.id === id);
  if (prod) {
    prod.isFavorite = !prod.isFavorite;
    saveData();
    renderApp();
  }
}

function deleteProduct(id) {
  if (confirm('Tem certeza que deseja remover este produto da sua planilha?')) {
    products = products.filter(p => p.id !== id);
    saveData();
    renderApp();
  }
}

function handleAddProduct(e) {
  e.preventDefault();

  const name = document.getElementById('inp-name').value;
  const category = document.getElementById('inp-category').value;
  const amzPrice = parseFloat(document.getElementById('inp-amz-price').value) || 0;
  const amzLink = document.getElementById('inp-amz-link').value;
  const sourcePrice = parseFloat(document.getElementById('inp-price').value) || 0;
  const freightEst = parseFloat(document.getElementById('inp-freight').value) || 0;
  const shpLink = document.getElementById('inp-link').value;
  const shpStore = document.getElementById('inp-store').value || 'Shopee';
  const priority = document.getElementById('inp-priority').value || 'Média';
  const notes = document.getElementById('inp-notes').value || '';

  const newProd = {
    id: 'prod-' + Date.now(),
    date: new Date().toLocaleDateString('pt-BR'),
    name,
    category,
    amzPrice,
    amzLink: amzLink || 'https://www.amazon.com.br',
    amzReviews: 'N/I',
    amzRating: 'N/I',
    amzComp: 1,
    shpStore,
    shpLink: shpLink || 'https://shopee.com.br',
    sourcePrice,
    freightEst,
    amazonTaxPct: 0.15,
    otherCosts: 0.00,
    isFavorite: false,
    priority,
    notes
  };

  products.unshift(newProd);
  saveData();
  closeModal('modal-add-product');
  document.getElementById('addProductForm').reset();
  renderApp();
}

// EDIÇÃO DE PRODUTOS
function openEditModal(id) {
  const prod = products.find(p => p.id === id);
  if (!prod) return;

  document.getElementById('edit-inp-id').value = prod.id;
  document.getElementById('edit-inp-name').value = prod.name || '';
  document.getElementById('edit-inp-category').value = prod.category || 'Cozinha / Organização';
  document.getElementById('edit-inp-amz-price').value = prod.amzPrice || 0;
  document.getElementById('edit-inp-amz-link').value = prod.amzLink || '';
  document.getElementById('edit-inp-price').value = prod.sourcePrice || 0;
  document.getElementById('edit-inp-freight').value = prod.freightEst || 0;
  document.getElementById('edit-inp-link').value = prod.shpLink || '';
  document.getElementById('edit-inp-store').value = prod.shpStore || '';
  document.getElementById('edit-inp-priority').value = prod.priority || 'Média';
  document.getElementById('edit-inp-notes').value = prod.notes || '';

  openModal('modal-edit-product');
}

function handleEditProduct(e) {
  e.preventDefault();

  const id = document.getElementById('edit-inp-id').value;
  const prod = products.find(p => p.id === id);
  if (!prod) return;

  prod.name = document.getElementById('edit-inp-name').value;
  prod.category = document.getElementById('edit-inp-category').value;
  prod.amzPrice = parseFloat(document.getElementById('edit-inp-amz-price').value) || 0;
  prod.amzLink = document.getElementById('edit-inp-amz-link').value;
  prod.sourcePrice = parseFloat(document.getElementById('edit-inp-price').value) || 0;
  prod.freightEst = parseFloat(document.getElementById('edit-inp-freight').value) || 0;
  prod.shpLink = document.getElementById('edit-inp-link').value;
  prod.shpStore = document.getElementById('edit-inp-store').value;
  prod.priority = document.getElementById('edit-inp-priority').value;
  prod.notes = document.getElementById('edit-inp-notes').value;

  saveData();
  closeModal('modal-edit-product');
  renderApp();
}

// UTILITÁRIOS DE MODAL
function openModal(modalId) {
  const m = document.getElementById(modalId);
  if (m) m.classList.add('active');
}

function closeModal(modalId) {
  const m = document.getElementById(modalId);
  if (m) m.classList.remove('active');
}

// EXPORTAÇÃO PARA CSV
function exportToCSV() {
  let csv = "Nº;Data da Pesquisa;Produto;Categoria;Preço Amazon (R$);Link Amazon;Avaliações;Nota;Concorrentes;Fornecedor Shopee;Link Shopee;Custo Unitário (R$);Frete Estimado (R$);Custo Entregue (R$);Taxa Amazon (15%);Custo Total Venda (R$);Lucro Líquido (R$);Margem (%);ROI (%)\n";

  products.forEach((p, idx) => {
    const fin = calculateFinancials(p);
    const row = [
      idx + 1,
      p.date || '10/09/2026',
      `"${(p.name || '').replace(/"/g, '""')}"`,
      `"${(p.category || '')}"`,
      p.amzPrice.toFixed(2).replace('.', ','),
      p.amzLink || '',
      p.amzReviews || 'N/I',
      p.amzRating || 'N/I',
      p.amzComp || 'N/I',
      `"${(p.shpStore || '')}"`,
      p.shpLink || '',
      p.sourcePrice.toFixed(2).replace('.', ','),
      fin.freightActual.toFixed(2).replace('.', ','),
      fin.costDelivered.toFixed(2).replace('.', ','),
      fin.taxAmazonR$.toFixed(2).replace('.', ','),
      fin.totalCostForSale.toFixed(2).replace('.', ','),
      fin.netProfit.toFixed(2).replace('.', ','),
      (fin.profitMarginPct * 100).toFixed(2).replace('.', ',') + '%',
      (fin.roiPct * 100).toFixed(2).replace('.', ',') + '%'
    ];
    csv += row.join(';') + "\n";
  });

  const blob = new Blob(["\ufeff" + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", `planilha_pesquisa_produtos_dropship_${new Date().toISOString().slice(0,10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// FORMATADORES
function formatBRL(val) {
  if (typeof val !== 'number' || isNaN(val)) return 'R$ 0,00';
  return val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function formatPct(val) {
  if (typeof val !== 'number' || isNaN(val)) return '0,00%';
  return (val * 100).toFixed(2).replace('.', ',') + '%';
}

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}
