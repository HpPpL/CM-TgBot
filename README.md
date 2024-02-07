# SearchFD

**Команда для сборки:**

```shell
docker build -t searchfd .
```

**Команда для подтягивания данных:**

```shell
docker run -d --mount type=bind,source=./Data.xlsx,target=/app/Data.xlsx --env-file .env searchfd
```