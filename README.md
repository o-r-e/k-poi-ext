# k-poi-ext
Some utility methods and variables for Apache POI (can shorten code)

# How to use with Maven

## Using snapshots

```xml
...
<repositories>
    <repository>
        <id>sonatype - snapshots</id>
        <url>https://s01.oss.sonatype.org/content/repositories/snapshots/</url>
        <snapshots>
            <enabled>true</enabled>
        </snapshots>
    </repository>
</repositories>
...
<dependency>
    <groupId>me.o-r-e</groupId>
    <artifactId>k-poi-ext</artifactId>
    <version>0.0.2-SNAPSHOT</version>
</dependency>
...
```

## Using releases

```xml
...
<dependency>
    <groupId>me.o-r-e</groupId>
    <artifactId>k-poi-ext</artifactId>
    <version>0.0.2</version>
</dependency>
...
```
