package utils;

public enum PricePattern {
    Bemin("#,### 원"), Coupang("#"), Kakao("#,###"), WeMake("#,###"), Yugio("#,###"), Naver("#");
    private final String value;
    PricePattern(String value) {this.value = value; }
    public String getValue()   { return value; }
}
