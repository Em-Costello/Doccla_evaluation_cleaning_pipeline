## Editing the mapping files

These files control how raw values from GP records get translated into standard categories. You are expected to update them — GP data is inconsistent, and new spellings, abbreviations, or terms will come up that the pipeline hasn't seen before.

**How to edit:**
1. Open the mapping file.
2. Each row has two columns: the raw value (exactly as it appears in the source data) on the left, and the category it should map to on the right.
3. To add a new value, add a new row with the raw value and its correct category. Save the CSV.

**Important — please read before editing:**
- **Spelling must match exactly.** `"Indian"` and `"Indian "` (with a trailing space) are treated as different values.
- **If a value in the data isn't in the mapping file, the pipeline will stop and show an error rather than guessing or skipping it.** This is intentional — it's there to make sure nothing gets mapped incorrectly or silently miscategorised. If you see an error like this, it usually means a new raw value needs to be added to the mapping file.
- Every value should map to exactly one category — don't map the same raw value to two different categories, or leave a category blank.

---

### Ethnicity mapping

Ethnicities have been mapped to broad categories from the [Gov.uk](https://www.ethnicity-facts-figures.service.gov.uk/style-guide/ethnic-groups/) webpage, based on the 2022 census categories.

#### Asian or Asian British
Indian
Pakistani
Bangladeshi
Chinese
Any other Asian background

#### Black, Black British, Caribbean or African
Caribbean
African
Any other Black, Black British, or Caribbean background

#### Mixed or multiple ethnic groups
White and Black Caribbean
White and Black African
White and Asian
Any other Mixed or multiple ethnic background

#### White
English, Welsh, Scottish, Northern Irish or British
Irish
Gypsy or Irish Traveller
Roma
Any other White background

#### Other ethnic group
Arab
Any other ethnic group