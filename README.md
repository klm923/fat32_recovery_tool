# undelete – FAT32 リカバリーツール

## 概要
FAT32 ドライブを直接読み取り、削除済みディレクトリエントリを走査して Excel に一覧化し、任意のファイルだけを安全に復元できる CLI ユーティリティです。Windows の `\\.\X:` デバイスパスを扱う低レイヤ I/O、FAT32 のメタデータ解析、Excel 連携、自動ディレクトリ生成／タイムスタンプ復元などを一つのスクリプトで実現しています。

## 主なスキル・技術ポイント
- **ディレクトリツリーの逆引き**：先頭クラスタから親ディレクトリを辿り、Excel の各行へフルパスを書き戻し。

```undelete.py
def lookup_path(excel_file_path: str):
    wb = openpyxl.load_workbook(excel_file_path)
    ws = wb.active
    parent_lookup = {}
    ...
    full_path = "\\".join(path_list)
    if full_path != "ROOT":
        print(f"復元パス: {full_path}")
    row[10].value = full_path
```

- **クラスタチェーン追跡と復元**：クラスタ番号を辿ってファイルを再構築し、Excel で指定した更新日時へ `os.utime` で揃えます。

```undelete.py
def salvage_file(excel_file_path: str):
    ...
    while file_size_rest > 0:
        current_cluster = get_next_cluster("D", current_cluster)
        cluster_chain.append(current_cluster)
        file_size_rest -= CLUSTER_SIZE
    ...
    with open(file_full_path, "wb") as out_f:
        out_f.write(file_data)
    update_datetime = datetime.strptime(row[7].value, "%Y-%m-%d %H:%M:%S")
    os.utime(path=file_full_path, times=(update_datetime.timestamp(), update_datetime.timestamp()))
```

- **FAT32 生データ解析**：ブートセクタからパラメータを抽出し、LFN(長いファイル名) や復旧対象拡張子フィルタを考慮しながら 32 バイトごとのエントリを解析。

```undelete.py
def read_raw_data(drive_letter: str, target_exts: List[str], xlsx_file: str):
    drive_path = f"\\\\.\\{drive_letter}:"
    ...
    boot_signature = raw_data[510:512]
    if boot_signature == b"\x55\xaa":
        BYTES_PER_SECTOR = unpack("<H", raw_data[11:13])[0]
        ...
        if not (extension_str in target_exts or full_filename_str[-6:] == '.pages' or ...):
            continue
        scan_results.append({
            "current_byte": current_byte,
            "filename": full_filename_str,
            ...
        })
    save_to_excel(scan_results, xlsx_file)
```

- **堅牢な CLI 設計**：`argparse` の排他的グループでスキャン／復元モードを切り替え、拡張子や Excel 出力先を柔軟に指定可能。

```undelete.py
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="FAT32ドライブからデータを復旧")
    parser.add_argument("--target_drive", "-t", type=str, required=True, ...)
    run_mode = parser.add_mutually_exclusive_group(required=True)
    run_mode.add_argument("--scan", "-s", action="store_true", ...)
    run_mode.add_argument("--restore", "-r", action="store_true", ...)
    parser.add_argument("--extensions", "-e", nargs="+", default=[...])
    parser.add_argument("--xlsx_file", "-x", type=str, default='fat32_scan_results.xlsx')
```

## 使い方
1. **スキャン**：指定ドライブを読み取って Excel に候補を書き出し、パス逆引きまで自動実行。
   ```
   python undelete.py --target_drive D --scan --extensions DOC XLS JPG --xlsx_file scan.xlsx
   ```
2. **Excel で復旧対象を選択**：`復旧チェック` 列に `1` を設定し、フルパスを確認。
3. **復元**：スキャン結果ファイルを読み、指定クラスタ連鎖から実データを再構築。
   ```
   python undelete.py --target_drive D --restore --xlsx_file scan.xlsx
   ```

## アピールポイント
- Windows のデバイスパスや FAT32 仕様に沿った低レイヤ実装により、一般的なファイル API では得られない情報を取得。
- LFN チェーンの正規化、Shift_JIS→UTF-16LE デコード、制御文字除去など日本語ファイル名の復元精度を高める工夫。
- Excel 連携を通じて、調査・復旧のワークフローを GUI で可視化。
- CLI オプションで拡張子フィルタや実行モードを切替でき、バッチ運用も容易。

