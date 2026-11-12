/* Chinese library classification; unknown legacy labels remain reviewable. */
const BOOK_CATEGORIES = ['000 總類','100 哲學類','200 宗教類','300 科學類（自然科學）','400 應用科學類','500 社會科學類','600 中國史地類','700 世界史地類','800 語言文學類','900 美術類（藝術類）'];
function normalizeBookCategory(value) {
  const raw=String(value??'').trim();
  if(!raw)return '';
  // A combined history/geography range cannot identify a single class.
  if(/^600\s*[-–—~～]\s*700/.test(raw))return raw;
  const numeric=raw.match(/^([0-9])[0-9]{2}(?:\.\d+)?(?:\s|$)/);
  if(numeric)return BOOK_CATEGORIES[Number(numeric[1])];
  const key=raw.replace(/[\s（）()]/g,'');
  const aliases=[['總類','總論','百科全書'],['哲學','哲學類','心理學','倫理學'],['宗教','宗教類'],['科學','科學類','自然','自然科學','自然科學類','科學類自然科學'],['應用科學','應用科學類'],['社會','社會科學','社會科學類'],['中國史地','中國史地類'],['世界史地','世界史地類'],['文學','語文','語言文學','語言文學類','文學類'],['藝術','藝術類','美術','美術類','美術類藝術類']];
  const index=aliases.findIndex(list=>list.includes(key));
  return index<0?raw:BOOK_CATEGORIES[index];
}
function bookCategoryBucket(value){const normalized=normalizeBookCategory(value);return BOOK_CATEGORIES.includes(normalized)?normalized:'待分類';}
function aggregateBookCategories(rows){const counts=Object.fromEntries(BOOK_CATEGORIES.map(label=>[label,0]));counts['待分類']=0;for(const row of rows||[]){const amount=Number(row.count);if(Number.isFinite(amount)&&amount>0)counts[bookCategoryBucket(row.category)]+=amount;}return Object.entries(counts).filter(([label,count])=>label!=='待分類'||count>0).map(([category,count])=>({category,count}));}
