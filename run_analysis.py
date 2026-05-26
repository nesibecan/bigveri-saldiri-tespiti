"""
UNSW-NB15 Saldırı Tespiti - Grafik Üretim Scripti
"""
import os
os.chdir(r"C:\Users\seyma\OneDrive\Masaüstü\bigdataproje")

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder, StandardScaler
from sklearn.metrics import (classification_report, confusion_matrix,
                             ConfusionMatrixDisplay, roc_auc_score, roc_curve,
                             accuracy_score, f1_score)
from sklearn.utils import class_weight
import warnings
warnings.filterwarnings('ignore')

plt.rcParams['figure.dpi'] = 150
sns.set_theme(style='whitegrid', palette='muted')

print("=== UNSW-NB15 Saldırı Tespiti Analizi ===\n")

# ── 1. VERİ YÜKLEME ──────────────────────────────────────────
print("1. Veri yükleniyor...")
train_df = pd.read_csv('data/UNSW_NB15_training-set.csv')
test_df  = pd.read_csv('data/UNSW_NB15_testing-set.csv')
print(f"   Eğitim seti: {train_df.shape}")
print(f"   Test seti  : {test_df.shape}")

# ── 2. SINIF DAĞILIMI ────────────────────────────────────────
print("\n2. Sınıf dağılımı grafikleri hazırlanıyor...")
label_counts  = train_df['label'].value_counts()
attack_counts = train_df['attack_cat'].value_counts()

fig, axes = plt.subplots(1, 2, figsize=(14, 5))
axes[0].pie(label_counts, labels=['Normal', 'Saldırı'], autopct='%1.1f%%',
            colors=['#4CAF50', '#F44336'], startangle=90,
            wedgeprops={'edgecolor': 'white', 'linewidth': 2})
axes[0].set_title('Eğitim Seti: Normal vs Saldırı', fontsize=13, fontweight='bold')

attack_only = attack_counts[attack_counts.index != 'Normal']
colors_bar  = sns.color_palette('Reds_r', len(attack_only))
axes[1].barh(attack_only.index, attack_only.values, color=colors_bar)
axes[1].set_xlabel('Kayıt Sayısı', fontsize=11)
axes[1].set_title('Saldırı Kategorileri', fontsize=13, fontweight='bold')
for i, v in enumerate(attack_only.values):
    axes[1].text(v + 30, i, str(v), va='center', fontsize=9)
plt.tight_layout()
plt.savefig('class_distribution.png', bbox_inches='tight')
plt.close()
print("   -> class_distribution.png kaydedildi")

# ── 3. PROTOKOL DAĞILIMI ─────────────────────────────────────
print("3. Protokol grafikleri hazırlanıyor...")
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
proto_counts = train_df['proto'].value_counts().head(10)
axes[0].bar(proto_counts.index, proto_counts.values,
            color=sns.color_palette('Blues_r', 10))
axes[0].set_title('En Yaygın 10 Protokol', fontsize=12, fontweight='bold')
axes[0].set_xlabel('Protokol')
axes[0].set_ylabel('Kayıt Sayısı')
axes[0].tick_params(axis='x', rotation=45)

state_counts = train_df['state'].value_counts().head(8)
axes[1].bar(state_counts.index, state_counts.values,
            color=sns.color_palette('Purples_r', 8))
axes[1].set_title('Bağlantı Durumu (State)', fontsize=12, fontweight='bold')
axes[1].set_xlabel('State')
axes[1].set_ylabel('Kayıt Sayısı')
plt.tight_layout()
plt.savefig('protocol_state.png', bbox_inches='tight')
plt.close()
print("   -> protocol_state.png kaydedildi")

# ── 4. ÖN İŞLEME ────────────────────────────────────────────
print("\n4. Veri ön işleme yapılıyor...")
cat_cols  = ['proto', 'service', 'state']
drop_cols = ['id', 'attack_cat']

def preprocess(df, encoders=None, fit=True):
    df = df.copy()
    df.drop(columns=drop_cols, inplace=True)
    if encoders is None:
        encoders = {}
    for col in cat_cols:
        if fit:
            le = LabelEncoder()
            df[col] = le.fit_transform(df[col].astype(str))
            encoders[col] = le
        else:
            le = encoders[col]
            df[col] = df[col].astype(str).map(
                lambda x, le=le: le.transform([x])[0] if x in le.classes_ else -1)
    df.fillna(df.median(numeric_only=True), inplace=True)
    return df, encoders

train_clean, encoders = preprocess(train_df, fit=True)
test_clean,  _        = preprocess(test_df, encoders=encoders, fit=False)

X_train = train_clean.drop(columns=['label'])
y_train = train_clean['label']
X_test  = test_clean.drop(columns=['label'])
y_test  = test_clean['label']

scaler = StandardScaler()
X_train_s = scaler.fit_transform(X_train)
X_test_s  = scaler.transform(X_test)
print(f"   X_train: {X_train_s.shape}, X_test: {X_test_s.shape}")

# ── 5. KORELASYON ISI HARİTASI ───────────────────────────────
print("\n5. Korelasyon ısı haritası hazırlanıyor...")
num_cols = [c for c in train_clean.select_dtypes(include=np.number).columns
            if c != 'label']
top_corr = train_clean[num_cols].corrwith(train_clean['label']).abs().nlargest(12).index.tolist()
corr_m   = train_clean[top_corr + ['label']].corr()

plt.figure(figsize=(11, 8))
mask = np.triu(np.ones_like(corr_m, dtype=bool))
sns.heatmap(corr_m, mask=mask, annot=True, fmt='.2f', cmap='coolwarm',
            center=0, square=True, linewidths=.5, cbar_kws={'shrink': .8})
plt.title('Label ile En Yüksek Korelasyonlu Özellikler\n(Korelasyon Matrisi)',
          fontsize=13, fontweight='bold')
plt.tight_layout()
plt.savefig('correlation_heatmap.png', bbox_inches='tight')
plt.close()
print("   -> correlation_heatmap.png kaydedildi")

# ── 6. MODEL EĞİTİMİ ────────────────────────────────────────
print("\n6. Random Forest modeli eğitiliyor...")
cw = class_weight.compute_class_weight(
    class_weight='balanced', classes=np.unique(y_train), y=y_train)
rf = RandomForestClassifier(n_estimators=100, max_depth=20,
                             min_samples_leaf=5,
                             class_weight={0: cw[0], 1: cw[1]},
                             random_state=42, n_jobs=-1)
rf.fit(X_train_s, y_train)
print("   Model eğitimi tamamlandı.")

# ── 7. DEĞERLENDİRME ─────────────────────────────────────────
print("\n7. Model değerlendirme grafikleri hazırlanıyor...")
y_pred  = rf.predict(X_test_s)
y_proba = rf.predict_proba(X_test_s)[:, 1]

report = classification_report(y_test, y_pred, target_names=['Normal', 'Saldırı'])
auc    = roc_auc_score(y_test, y_proba)
acc    = accuracy_score(y_test, y_pred)
f1     = f1_score(y_test, y_pred)
cm     = confusion_matrix(y_test, y_pred)
tn, fp, fn, tp = cm.ravel()

print("\n" + "="*60)
print("       SINIFLANDIRMA RAPORU")
print("="*60)
print(report)
print(f"ROC-AUC : {auc:.4f}")
print(f"Accuracy: {acc*100:.2f}%  |  F1: {f1:.4f}")
print(f"TP={tp:,}  FP={fp:,}  FN={fn:,}  TN={tn:,}")

# Confusion Matrix
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
ConfusionMatrixDisplay(confusion_matrix=cm,
                       display_labels=['Normal','Saldırı']).plot(
    ax=axes[0], colorbar=False, cmap='Blues')
axes[0].set_title('Karmaşıklık Matrisi (Ham)', fontsize=12, fontweight='bold')

cm_n = cm.astype('float') / cm.sum(axis=1)[:, np.newaxis]
ConfusionMatrixDisplay(confusion_matrix=cm_n.round(3),
                       display_labels=['Normal','Saldırı']).plot(
    ax=axes[1], colorbar=False, cmap='Greens')
axes[1].set_title('Karmaşıklık Matrisi (Normalize)', fontsize=12, fontweight='bold')
plt.tight_layout()
plt.savefig('confusion_matrix.png', bbox_inches='tight')
plt.close()
print("   -> confusion_matrix.png kaydedildi")

# ROC Eğrisi
fpr, tpr, _ = roc_curve(y_test, y_proba)
plt.figure(figsize=(7, 5))
plt.plot(fpr, tpr, color='#1565C0', lw=2, label=f'ROC (AUC={auc:.4f})')
plt.plot([0,1],[0,1],'k--', lw=1, label='Rastgele')
plt.fill_between(fpr, tpr, alpha=0.15, color='#1565C0')
plt.xlabel('Yanlış Pozitif Oranı (FPR)', fontsize=11)
plt.ylabel('Doğru Pozitif Oranı (TPR)', fontsize=11)
plt.title('ROC Eğrisi – Random Forest', fontsize=13, fontweight='bold')
plt.legend(loc='lower right', fontsize=10)
plt.tight_layout()
plt.savefig('roc_curve.png', bbox_inches='tight')
plt.close()
print("   -> roc_curve.png kaydedildi")

# Feature Importance
feat_imp = pd.Series(rf.feature_importances_, index=X_train.columns)
top20    = feat_imp.nlargest(20)
plt.figure(figsize=(10, 7))
top20.sort_values().plot(kind='barh', color=sns.color_palette('YlOrRd_r', 20))
plt.xlabel('Önem Skoru', fontsize=11)
plt.title('En Önemli 20 Özellik (Feature Importance)', fontsize=13, fontweight='bold')
plt.tight_layout()
plt.savefig('feature_importance.png', bbox_inches='tight')
plt.close()
print("   -> feature_importance.png kaydedildi")

# ── 8. RİSK ANALİZİ ─────────────────────────────────────────
print("\n8. Risk analizi grafiği hazırlanıyor...")
fig, ax = plt.subplots(figsize=(11, 5))
cats   = ['Doğru Normal\n(TN)', 'Yanlış Pozitif\n(FP)\n[ETİK RİSK]',
          'Yanlış Negatif\n(FN)\n[GÜVENLİK RİSKİ]', 'Doğru Saldırı\n(TP)']
vals   = [tn, fp, fn, tp]
bclrs  = ['#4CAF50', '#FF9800', '#F44336', '#2196F3']
bars   = ax.bar(cats, vals, color=bclrs, edgecolor='white', linewidth=1.5)
for bar, val in zip(bars, vals):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 50,
            f'{val:,}', ha='center', va='bottom', fontsize=10, fontweight='bold')
ax.set_title('Sınıflandırma Sonuçları ve Risk Kategorileri',
             fontsize=13, fontweight='bold')
ax.set_ylabel('Kayıt Sayısı', fontsize=11)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))
plt.tight_layout()
plt.savefig('risk_analysis.png', bbox_inches='tight')
plt.close()
print("   -> risk_analysis.png kaydedildi")

# ── ÖZET KAYDET ──────────────────────────────────────────────
results = {
    'acc': float(acc), 'f1': float(f1), 'auc': float(auc),
    'tp': int(tp), 'fp': int(fp), 'fn': int(fn), 'tn': int(tn),
    'total_train': int(len(y_train)), 'total_test': int(len(y_test)),
    'normal_train': int((y_train==0).sum()), 'attack_train': int((y_train==1).sum()),
    'fp_rate': float(fp/(fp+tn)*100), 'fn_rate': float(fn/(fn+tp)*100),
    'recall': float(tp/(tp+fn)*100), 'precision': float(tp/(tp+fp)*100),
    'top5_features': list(feat_imp.nlargest(5).index)
}
import json
with open('results.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)

print("\n" + "="*60)
print("TÜM GRAFİKLER VE SONUÇLAR BAŞARIYLA OLUŞTURULDU!")
print("="*60)
print(f"\nÖZET:")
print(f"  Doğruluk : %{acc*100:.2f}")
print(f"  F1 Skoru : {f1:.4f}")
print(f"  ROC-AUC  : {auc:.4f}")
print(f"  FP (Etik riski)     : {fp:,}")
print(f"  FN (Güvenlik riski) : {fn:,}")
