# -*- coding: utf-8 -*-
"""sample_fixed.html の単一 hand-wobble を、大きさの帯ごとの4本に置き換える。"""
import io
import sys

PATH = r"C:\Users\topge\OneDrive\ドキュメント\GitHub\claude_playground\work\graphreco-tegaki\sample_fixed.html"

with io.open(PATH, encoding="utf-8") as f:
    s = f.read()

OLD_DEF = '''  <filter id="hand-wobble" x="-30%" y="-30%" width="160%" height="160%">
    <feTurbulence type="fractalNoise" baseFrequency="0.02 0.03" numOctaves="2" seed="7" result="noise"/>'''

NEW_DEF = '''  <!-- 揺らぎフィルタは「対象が置かれた座標系の単位」で効くため、大きさの帯ごとに4本用意する -->
  <filter id="hand-wobble-xs" x="-40%" y="-40%" width="180%" height="180%">
    <feTurbulence type="fractalNoise" baseFrequency="0.25 0.31" numOctaves="2" seed="7" result="noise"/>
    <feDisplacementMap in="SourceGraphic" in2="noise" scale="0.67" xChannelSelector="R" yChannelSelector="G"/>
  </filter>
  <filter id="hand-wobble-s" x="-30%" y="-30%" width="160%" height="160%">
    <feTurbulence type="fractalNoise" baseFrequency="0.055 0.07" numOctaves="2" seed="7" result="noise"/>
    <feDisplacementMap in="SourceGraphic" in2="noise" scale="3.0" xChannelSelector="R" yChannelSelector="G"/>
  </filter>
  <filter id="hand-wobble-m" x="-15%" y="-15%" width="130%" height="130%">
    <feTurbulence type="fractalNoise" baseFrequency="0.010 0.013" numOctaves="1" seed="7" result="noise"/>
    <feDisplacementMap in="SourceGraphic" in2="noise" scale="3.0" xChannelSelector="R" yChannelSelector="G"/>
  </filter>
  <filter id="hand-wobble-l" x="-5%" y="-15%" width="110%" height="130%">
    <feTurbulence type="fractalNoise" baseFrequency="0.005 0.007" numOctaves="1" seed="7" result="noise"/>'''

# 置換は「一意に決まる文字列 -> 置換後」の順序付きリスト。件数を数え、想定と違えば異常終了する。
subs = [
    (OLD_DEF, NEW_DEF, 1),
    # 旧hand-arrow（viewBox 60x24）を長辺80単位に揃え直す。これで小帯がそのまま効く
    ('''  <symbol id="hand-arrow" viewBox="0 0 60 24">
    <g filter="url(#hand-wobble)">
      <path d="M2 12 Q30 6 48 12" fill="none" stroke="#2A2723" stroke-width="2.8" stroke-linecap="round"/>
      <path d="M48 12 L40 7 M48 12 L40 17" fill="none" stroke="#2A2723" stroke-width="2.2" stroke-linecap="round"/>''',
     '''  <symbol id="hand-arrow" viewBox="0 0 80 32">
    <g filter="url(#hand-wobble-s)">
      <path d="M3 16 Q40 8 64 16" fill="none" stroke="#2A2723" stroke-width="3.2" stroke-linecap="round"/>
      <path d="M64 16 L53 9 M64 16 L53 23" fill="none" stroke="#2A2723" stroke-width="2.6" stroke-linecap="round"/>''', 1),
    # check-mark は viewBox 18単位。極小帯
    ('''  <symbol id="check-mark" viewBox="0 0 18 16">
    <g filter="url(#hand-wobble)">''',
     '''  <symbol id="check-mark" viewBox="0 0 18 16">
    <g filter="url(#hand-wobble-xs)">''', 1),
    # 散布図は viewBox 420単位のCSSピクセル相当。中帯
    ('''        <svg viewBox="0 0 420 320" width="100%" style="max-width:460px;">
          <!-- 線画・図形はgでまとめて揺らぎフィルタを適用。text要素はgの外に置き可読性を保つ -->
          <g filter="url(#hand-wobble)">''',
     '''        <svg viewBox="0 0 420 320" width="100%" style="max-width:460px;">
          <!-- 線画・図形はgでまとめて揺らぎフィルタを適用。text要素はgの外に置き可読性を保つ。
               viewBoxの一辺が420単位なので中帯を当てる -->
          <g filter="url(#hand-wobble-m)">''', 1),
    # .card::before は幅900px級 -> 大帯 / .sticky-note::before は300px級 -> 中帯
    ('''    border-radius: 16px 20px 15px 22px / 20px 15px 22px 16px;
    box-shadow: 0 6px 14px var(--shadow);
    filter:url(#hand-wobble);''',
     '''    border-radius: 16px 20px 15px 22px / 20px 15px 22px 16px;
    box-shadow: 0 6px 14px var(--shadow);
    filter:url(#hand-wobble-l);''', 1),
    ('''    background:var(--highlight); border-radius:2px 5px 3px 6px;
    box-shadow:3px 5px 9px var(--shadow);
    filter:url(#hand-wobble);''',
     '''    background:var(--highlight); border-radius:2px 5px 3px 6px;
    box-shadow:3px 5px 9px var(--shadow);
    filter:url(#hand-wobble-m);''', 1),
]

for old, new, want in subs:
    got = s.count(old)
    if got != want:
        sys.exit("想定と一致しません: %d件 (期待 %d件) / 先頭: %s" % (got, want, old[:60]))
    s = s.replace(old, new)

# 残る `url(#hand-wobble)` は viewBox 80〜100単位のシンボル。すべて小帯へ。
rest = s.count('url(#hand-wobble)')
s = s.replace('filter="url(#hand-wobble)"', 'filter="url(#hand-wobble-s)"')

leftover = s.count('url(#hand-wobble)')
if leftover:
    sys.exit("未変換の hand-wobble が %d件 残りました" % leftover)

with io.open(PATH, "w", encoding="utf-8") as f:
    f.write(s)

print("OK: 個別置換 %d件 / 小帯へ一括 %d件" % (len(subs), rest))
