FROM openjdk:17
ADD ./DocHouse-0.0.1-SNAPSHOT.jar DocHouse-0.0.1-SNAPSHOT.jar
ENTRYPOINT ["java","-jar","DocHouse-0.0.1-SNAPSHOT.jar"]