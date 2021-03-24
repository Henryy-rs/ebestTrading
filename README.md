# ebestTrading


 wikidocs문서 알고리즘 트레이딩(https://wikidocs.net/6316)를 참고하여 개발중인
ebest xing API를 이용한 트레이딩 프로그램이다.

xingACE를 설치하고, DevCenter에서 모든 Res파일을 다운받으면 사용 가능하다.

한글 폰트를 설정하지 않으면 pie graph에 한글이 깨진다.
폰트 다운 후 설치된 matplotlib의 font폴더에 복사하고 matplotlib cashe폴더를 삭제한 후 실행하면 된다.


# 현재 작동하는 기능

- 주문(시간외종가, 시간외단일가 포함)
- 잔고조회
- 보유주식조회
- 주문조회
- 주문정정
- 주문취소
- 주가정보조회(네이버증권 웹으로 대체)
- 포트폴리오그래프
- 목표 수익률 자동매도(v2.0)
- ~~자동매매~~

기타 기능
- 보유주식목록이나 주문목록에서 tableWidget을 클릭하면 종목번호 혹은 주문번호를 자동으로 입력됨.
- 위의 방법으로 (목표 수익률)자동매도 리스트에 포함시킬 수 있음 
- 주문시 정상적으로 들어가면 statusbar에 초록불이 들어오고 아니면 빨간불이 들어옴

![image](https://user-images.githubusercontent.com/28619620/112286970-388fd000-8ccf-11eb-8d81-668d46f63a7a.png)
![image](https://user-images.githubusercontent.com/28619620/112286670-e484eb80-8cce-11eb-9104-6215a2270604.png)
![image](https://user-images.githubusercontent.com/28619620/112286555-c7e8b380-8cce-11eb-9dbe-c9d43991da1b.png)
![image](https://user-images.githubusercontent.com/28619620/112286383-9cfe5f80-8cce-11eb-9bbb-989a06e99c2e.png)
