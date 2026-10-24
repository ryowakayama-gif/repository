#!/bin/sh
# 全成果品をビルドする。スクラップパッドに style/ と build/ を複写したうえで、そこで実行する。
#   sh build_all.sh
set -e
python3 figs.py >/dev/null 2>&1
python3 figs_soan.py >/dev/null 2>&1
for f in ch_*.py; do
  n=${f#ch_}; n=${n%.py}
  case $n in
    kohyo) { cat style_kohyo.py; cat fmt.py; echo; cat $f; } > mkf_$n.py ;;
    iinkai2|shidai2|qa2) { cat style_iinkai1.py; cat fmt.py; echo; cat fmt_iinkai1.py; echo; cat $f; } > mkf_$n.py ;;
    *) { cat fmt.py; echo; cat $f; } > mkf_$n.py ;;
  esac
  python3 mkf_$n.py >/dev/null 2>&1 || echo "FAIL $n"
done
python3 mk_toc.py "out/計画素案_第10期.docx" mkf_soan.py >/dev/null 2>&1
python3 mk_toc.py "out/計画素案_第10期_公表用.docx" mkf_kohyo.py "out/_toc_pages_kohyo.json" >/dev/null 2>&1
echo BUILD_DONE
