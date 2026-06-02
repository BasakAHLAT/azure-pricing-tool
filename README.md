# 🔵 Azure Pricing Tool

> Azure kaynaklarının Pay-as-you-go ve rezervasyon fiyatlarını otomatik hesaplar, Excel raporu üretir.

Teknik presales süreçlerinde müşteri teklifleri için dakikalar içinde maliyet analizi yapmanı sağlar.

---

## Ne İşe Yarar?

`config.json` dosyasına VM SKU'larını, disk tiplerini, bölgeyi ve diğer Azure kaynaklarını tanımlarsın. Script Microsoft'un resmi fiyat API'sine bağlanarak:

- **Pay-as-you-go** fiyatını çeker
- **1 yıllık rezervasyon** fiyatını çeker
- **3 yıllık rezervasyon** fiyatını çeker
- Tasarruf yüzdelerini hesaplar
- Terminal'de renkli tablo gösterir
- Excel raporu üretir

---

## Kurulum

```bash
git clone https://github.com/BasakAHLAT/azure-pricing-tool
cd azure-pricing-tool
pip install -r requirements.txt
```

### 1. config.json'u Düzenle

```json
{
  "region": "westeurope",
  "currency": "USD",
  "hours_per_month": 730,
  "resources": {
    "vms": [
      {"sku": "Standard_D2s_v5", "count": 1}
    ],
    "disks": [
      {"type": "P10", "count": 2}
    ],
    "storage_accounts": [
      {"type": "LRS", "gb": 500}
    ]
  }
}
```

### 2. Çalıştır

```bash
python azure_pricing.py
```

### 3. Çıktılar

| Çıktı | Açıklama |
|-------|----------|
| Terminal | Servis tipine göre gruplanmış renkli tablo |
| `azure_pricing_report.xlsx` | PAYG + rezervasyon karşılaştırması, her servis ayrı sheet |

---

## Desteklenen Kaynaklar

| Kaynak | config.json Anahtarı | Örnek |
|--------|----------------------|-------|
| Virtual Machines | `vms` | `Standard_D2s_v5` |
| Managed Disks | `disks` | `P10`, `P20`, `P30` |
| App Service | `app_services` | `P1v3`, `P2v3` |
| SQL Database | `sql_databases` | `GP_Gen5_2` |
| Storage Account | `storage_accounts` | `LRS`, `GRS`, `ZRS` |

---

## Desteklenen Bölgeler

`westeurope`, `northeurope`, `eastus`, `westus2`, `uksouth`, `germanywestcentral` ve tüm Azure bölgeleri.

---

## Veri Kaynağı

Microsoft'un resmi [Azure Retail Prices API](https://prices.azure.com/api/retail/prices) — auth gerektirmez, her zaman güncel fiyatları döner.

---

## Geliştiren

**Başak Ahlat** 
