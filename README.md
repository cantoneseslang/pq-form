vercel_minimal/pq-form 配下一式が GitHub（cantoneseslang/pq-form）に push され、Vercel にデプロイされています
［https://github.com/cantoneseslang/pq-form］。
変更するならこのディレクトリ内の index.html・app.js・styles.css・api/pq_form/* などを編集してください。

## 原材料 QC（測厚計 OCR）

モバイル向けページ: `/material-qc/`

測厚計 LCD 写真から Gemini Vision で基材（鉄）と塗装厚（μm）を読取り、Google Sheets + Drive に保存します。

### Vercel 環境変数（追加）

| 変数 | 必須 | 説明 |
|------|------|------|
| `GEMINI_API_KEY` | はい | [Google AI Studio](https://aistudio.google.com/apikey) の API キー |
| `RAW_MATERIAL_QC_DRIVE_FOLDER_ID` | いいえ | 写真の Drive 保存（任意・Supabase 失敗時のフォールバック） |
| `RAW_MATERIAL_QC_STORAGE_BUCKET` | いいえ | Supabase Storage バケット名（省略時 `material-qc`） |
| `GOOGLE_DRIVE_DELEGATED_USER` | いいえ | Drive 保存時の Workspace ユーザー（ドメイン委任） |
| `RAW_MATERIAL_QC_SHEET_ID` | いいえ | 記録先スプレッドシート ID（省略時 `PQFORM_SHEET_ID` = `1u_fsEVAumMySLx8fZdMP5M4jgHiGG6ncPjFEXSXHQ1M`） |
| `RAW_MATERIAL_QC_SHEET_NAME` | いいえ | タブ名（省略時 `RAW_MATERIAL_QC`） |
| `GEMINI_MODEL` | いいえ | 省略時 `gemini-2.5-flash` |

### 事前設定

1. 写真は **Supabase Storage**（`material-qc` バケット）に保存し、公開 URL を Sheets に記録
2. （任意）Drive フォルダへも保存したい場合は `RAW_MATERIAL_QC_DRIVE_FOLDER_ID` を設定
3. 初回 submit 時に `RAW_MATERIAL_QC` タブが自動作成されます（[PQ-Form スプレッドシート](https://docs.google.com/spreadsheets/d/1u_fsEVAumMySLx8fZdMP5M4jgHiGG6ncPjFEXSXHQ1M/edit)）

