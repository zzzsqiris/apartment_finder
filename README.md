# Apartment Finder

这个 repo 是一个学校、地址或坐标周边的租房筛选流水线：

1. 用 Google Places 或免费 OSM 找学校附近的小区。
2. 清洗结果，剔除明显不是公寓/住宅的地点。
3. 只保留带网站的选项，自动进入官网/房源页抽取价格、面积、洗衣、停车等信息。
4. 输出 Markdown、CSV、Excel、Word，方便继续人工筛选。

## Install

```bash
python3 -m pip install -r requirements.txt
```

## Quick Start: General Search

交互式通用版：

```bash
./run_apartment_finder.sh
```

它会依次问你：

- 学校、地址或地点名
- 可选坐标。如果不填，会用 Google Places 找这个地点
- 搜索半径，比如 `2200`
- 租金预算，比如 `2500`
- 本次 Google API 预算上限，比如 `5`
- 是否跳过路线计算来省钱
- 可选商圈关键词，比如 `Westwood, Wilshire Corridor`
- 可选必须检查的小区名，比如 `Lindbrook Manor`

输出文件会用地点名自动命名。例如搜索 `University of Southern California` 时，文件名会类似：

- `university_of_southern_california_apartments_google.md`
- `university_of_southern_california_apartments_clean.csv`
- `university_of_southern_california_screening.xlsx`
- `university_of_southern_california_screening_summary.docx`

## Example: UCLA

这个 repo 里保留了一个 UCLA 快捷脚本，用来复现最初的搜索：

```bash
./run_ucla_google_budget.sh
```

## Manual Step 1: Search Apartments

同样的逻辑也可以手动写成任意地点，例如下面是 UCLA：

```bash
python3 find_apartments_google_budget.py \
  --school "UCLA" \
  --center-lat 34.0703 \
  --center-lon -118.4448 \
  --radius-meters 2200 \
  --max-walk-minutes 20 \
  --budget-usd 5 \
  --osm-markdown apartments_osm.md \
  --seed "Lindbrook Manor" \
  --area "Westwood" \
  --area "Westwood Village" \
  --area "Wilshire Corridor" \
  --skip-routes \
  --output apartments_google_budget.md
```

免费 OSM 版：

```bash
python3 find_apartments_osm.py \
  --school "UCLA" \
  --radius-meters 2200 \
  --max-walk-minutes 20 \
  --output apartments_osm.md
```

## Step 2: Clean Results

```bash
python3 clean_google_budget_results.py \
  --input apartments_google_budget.csv \
  --output-md apartments_google_budget_clean.md \
  --output-csv apartments_google_budget_clean.csv
```

## Step 3: Screen Websites

这个步骤会自动跳过没有网站的房源，然后抓官网和相关子页面，比如 floor plans、availability、amenities、FAQ。

```bash
python3 screen_apartments_from_web.py \
  --input apartments_google_budget_clean.csv \
  --output-csv apartment_screening_auto.csv \
  --output-xlsx apartment_screening_auto.xlsx \
  --output-docx apartment_screening_auto_summary.docx
```

或者直接：

```bash
./run_website_screening.sh
```

输出：

- `apartment_screening_auto.csv`
- `apartment_screening_auto.xlsx`
- `apartment_screening_auto_summary.docx`

## Manual Overrides

很多公寓网站会把价格藏在动态组件里，或者需要打电话才给准确价格。确认过的信息可以写进 `manual_overrides.json`，下次运行时会覆盖自动抽取结果。

现在已经放入两个决赛圈房源：

- Midvale Court Apartments Westwood
- Lindbrook Manor

## Cost Notes

- `screen_apartments_from_web.py` 不使用 Google API，不花 Google Maps 钱。
- Google Places 搜索脚本会缓存结果，重复运行会尽量复用 `.google_budget_cache.json`。
- `--skip-routes` 会跳过 Routes API，用直线距离排序来省钱。

## Limitations

自动网页抽取适合初筛，不适合替代最终确认。以下情况需要人工复核：

- 价格通过 JavaScript 动态加载。
- 网站只写 `Call for pricing`。
- 房型表在图片或 PDF 里。
- 洗衣/停车信息写得很模糊，比如 `select units`。
- 第三方 listing 和官网价格不一致。
