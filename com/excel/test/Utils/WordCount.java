package com.excel.test.Utils;

/**
 * Created by smt2 on 16-12-21.
 */
public class WordCount {


    private int wordCount(String test) {
        String symbol = "!，。、()（）";
        int chineseCount = CharMatcher.inRange('一', '龥').retainFrom(test).length();
        String symbolAndEnglish = CharMatcher.inRange('一', '龥').or(CharMatcher.anyOf("\n")).replaceFrom(test, " ");
        int symbolCount = CharMatcher.anyOf(symbol).retainFrom(symbolAndEnglish).length();
        Iterable<String> splitter = Splitter.on(' ')
                .trimResults()
                .omitEmptyStrings()
                .split(CharMatcher.anyOf(symbol).replaceFrom(symbolAndEnglish, " "));
        int englishCount = Lists.newArrayList(splitter).size();
        return chineseCount + englishCount + symbolCount;
    }
    2016年11月22日星期二

}
