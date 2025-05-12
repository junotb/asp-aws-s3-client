# cAwsS3Client.asp

Classic ASP 환경에서 AWS Signature Version 4 방식으로 S3 객체를 안전하게 가져오는 클라이언트 클래스입니다.

이 클래스는 VTT, 이미지, JSON 등의 파일을 S3에서 직접 요청해 가져오고, 적절한 Content-Type으로 응답하도록 지원합니다.

---

## 파일 구성

### `cAwsS3Client.asp`

- AWS S3 요청을 위한 클래스로 다음 기능을 제공합니다:
  - `GetS3ObjectContent(key)` : S3 객체 내용 가져오기
  - `GetContentTypeFromKey(key)` : 확장자에 따른 Content-Type 추론
  - 내부적으로 AWS Signature V4 헤더 자동 생성 및 인증 처리

### `example.asp`

- HTTP GET 파라미터 `S3ObjectKey`를 기준으로 지정된 S3 객체를 가져와 브라우저로 전송합니다.
- 예:
  ```http
  GET /example.asp?S3ObjectKey=/zoom-data/vtt/123/456.vtt

## 설정 가이드

Private S3                  ' 예시: s3
Private S3_BUCKET           ' 예시: my-bucket
Private S3_REGION           ' 예시: us-east-1
Private S3_DOMAIN           ' 예시: amazonaws.com
Private S3_ROOT_DIRECTORY   ' 예시: my-directory
Private S3_ACL              ' 예시: public-read
Private S3_USER_NAME        ' 예시: my-user
Private S3_ACCESS_KEY       ' 예시: my-access-key
Private S3_SECRET_KEY       ' 예시: my-secret-key
