# ebestTrading


 wikidocs문서 알고리즘 트레이딩(https://wikidocs.net/6316)를 참고하여 개발중인
ebest xing API를 이용한 트레이딩 프로그램이다.

xingACE를 설치하고, DevCenter에서 모든 Res파일을 다운받으면 사용 가능하다.

한글 폰트를 설정하지 않으면 pie graph에 한글이 깨진다.
폰트 다운 후 설치된 matplotlib의 font폴더에 복사하고 matplotlib cashe폴더를 삭제한 후 실행하면 된다.


# 현재 작동하는 기능

- 매수
- 매도
- 잔고조회
- 보유주식조회
- 주문조회
- 주문정정
- 주문취소
- 주가정보조회(네이버증권 웹으로 대체)
- 포트폴리오그래프
- ~~자동매매~~

기타 편의 기능
- 보유주식목록이나 주문목록에서 tableWidget을 클릭하면 종목번호 혹은 주문번호를 자동으로 입력해줌
- 주문시 정상적으로 들어가면 statusbar에 초록불이 들어오고 아니면 빨간불이 들어옴

![balance](https://user-images.githubusercontent.com/28619620/107951843-4e2d1e00-6fdc-11eb-87f4-e03fe31e4c62.png)
![order_list](https://user-images.githubusercontent.com/28619620/107951908-656c0b80-6fdc-11eb-84ae-4a8448b03dbe.png)
![search](https://user-images.githubusercontent.com/28619620/107951986-7e74bc80-6fdc-11eb-8066-c60d2c2dc6a0.png)
![porfolio](https://user-images.githubusercontent.com/28619620/107951955-7583eb00-6fdc-11eb-8eba-96cb39a00e3d.png)
