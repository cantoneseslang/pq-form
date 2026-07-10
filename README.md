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
| `RAW_MATERIAL_QC_DRIVE_FOLDER_ID` | はい | **共享雲端硬碟**（Shared Drive）内フォルダ ID |
| `RAW_MATERIAL_QC_STORAGE_BUCKET` | いいえ | Supabase フォールバック用バケット名（省略時 `material-qc`） |
| `GOOGLE_DRIVE_DELEGATED_USER` | いいえ | （通常不要）Shared Drive 利用時は未設定で OK |
| `RAW_MATERIAL_QC_SHEET_ID` | いいえ | 記録先スプレッドシート ID（省略時 `PQFORM_SHEET_ID` = `1u_fsEVAumMySLx8fZdMP5M4jgHiGG6ncPjFEXSXHQ1M`） |
| `RAW_MATERIAL_QC_SHEET_NAME` | いいえ | タブ名（省略時 `RAW_MATERIAL_QC`） |
| `GEMINI_MODEL` | いいえ | 省略時 `gemini-2.5-flash` |

### 事前設定（相片 → Google Drive のみ）

1. Google Drive にフォルダを作成（例: `material-qc-photos`）
2. Vercel に設定:
   - `RAW_MATERIAL_QC_DRIVE_FOLDER_ID` = フォルダ ID
   - `GOOGLE_DRIVE_DELEGATED_USER` = そのフォルダの所有者メール
3. Google Admin で Domain-wide Delegation を有効化（サービスアカウント + Drive スコープ）
4. **Supabase には保存しません**

