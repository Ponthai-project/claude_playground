#!/bin/sh
# BOM 付き UTF-8 への合成スクリプト
#
# なぜ必要か：
#   Write ツールも一般的なエディタ設定も UTF-8 BOM を付けない。しかし Windows
#   PowerShell 5.1 は BOM の無い .ps1 を CP932 として読むため、日本語リテラル
#   （deny 理由文・パターン名）を含む本フックは BOM を落とすと本体が壊れる。
#   そこで BOM 3 バイト（EF BB BF）を先頭に置いてから本文を連結する。
#
# 使い方：  sh build-bom.sh <入力(BOMなし)> <出力(BOM付き)>
#
# 注意：出力先を毎回変えないこと。試行錯誤では同名で上書きする（許可札の使い捨てを防ぐため）。

set -e

SRC="$1"
DST="$2"

if [ -z "$SRC" ] || [ -z "$DST" ]; then
  echo "usage: sh build-bom.sh <src-without-bom> <dst-with-bom>" >&2
  exit 1
fi

if [ ! -f "$SRC" ]; then
  echo "ERROR: 入力ファイルが無い: $SRC" >&2
  exit 1
fi

# 入力が既に BOM 付きなら二重付与になるので拒否する。
head -c 3 "$SRC" | od -A n -t x1 | tr -d ' \n' | grep -q '^efbbbf$' && {
  echo "ERROR: 入力に既に BOM がある: $SRC" >&2
  exit 1
}

printf '\357\273\277' > "$DST"
cat "$SRC" >> "$DST"

echo "built: $DST"
head -c 3 "$DST" | od -A n -t x1
wc -c < "$DST"
