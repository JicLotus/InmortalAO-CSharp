CREATE DATABASE  IF NOT EXISTS `inmortalao` /*!40100 DEFAULT CHARACTER SET latin1 */;
USE `inmortalao`;
-- MySQL dump 10.13  Distrib 5.7.17, for Win64 (x86_64)
--
-- Host: localhost    Database: inmortalao
-- ------------------------------------------------------
-- Server version	5.7.18-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `charatrib`
--

DROP TABLE IF EXISTS `charatrib`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charatrib` (
  `IndexPJ` int(11) NOT NULL DEFAULT '0',
  `AT1` int(11) DEFAULT '0',
  `AT2` int(11) DEFAULT '0',
  `AT3` int(11) DEFAULT '0',
  `AT4` int(11) DEFAULT '0',
  `AT5` int(11) DEFAULT '0',
  PRIMARY KEY (`IndexPJ`),
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charatrib`
--

LOCK TABLES `charatrib` WRITE;
/*!40000 ALTER TABLE `charatrib` DISABLE KEYS */;
INSERT INTO `charatrib` VALUES (944,19,19,18,6,13),(945,18,14,18,15,9),(946,19,18,16,13,8),(947,11,11,11,10,12),(948,11,11,11,10,12),(949,11,11,11,10,12),(950,11,11,11,10,12),(951,11,11,11,10,12),(952,11,11,11,10,12),(953,11,11,11,10,12),(954,11,11,11,10,12),(955,11,11,11,10,12),(956,NULL,NULL,NULL,NULL,NULL),(957,NULL,NULL,NULL,NULL,NULL),(958,NULL,NULL,NULL,NULL,NULL),(959,NULL,NULL,NULL,NULL,NULL),(960,19,19,19,10,8),(0,18,18,18,10,6),(962,17,15,15,17,6),(963,0,0,0,0,0),(964,0,0,0,0,0),(965,18,17,16,13,6);
/*!40000 ALTER TABLE `charatrib` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charbanco`
--

DROP TABLE IF EXISTS `charbanco`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charbanco` (
  `IndexPJ` int(11) NOT NULL DEFAULT '0',
  `OBJ1` int(11) DEFAULT '0',
  `CANT1` int(11) DEFAULT '0',
  `OBJ2` int(11) DEFAULT '0',
  `CANT2` int(11) DEFAULT '0',
  `OBJ3` int(11) DEFAULT '0',
  `CANT3` int(11) DEFAULT '0',
  `OBJ4` int(11) DEFAULT '0',
  `CANT4` int(11) DEFAULT '0',
  `OBJ5` int(11) DEFAULT '0',
  `CANT5` int(11) DEFAULT '0',
  `OBJ6` int(11) DEFAULT '0',
  `CANT6` int(11) DEFAULT '0',
  `OBJ7` int(11) DEFAULT '0',
  `CANT7` int(11) DEFAULT '0',
  `OBJ8` int(11) DEFAULT '0',
  `CANT8` int(11) DEFAULT '0',
  `OBJ9` int(11) DEFAULT '0',
  `CANT9` int(11) DEFAULT '0',
  `OBJ10` int(11) DEFAULT '0',
  `CANT10` int(11) DEFAULT '0',
  `OBJ11` int(11) DEFAULT '0',
  `CANT11` int(11) DEFAULT '0',
  `OBJ12` int(11) DEFAULT '0',
  `CANT12` int(11) DEFAULT '0',
  `OBJ13` int(11) DEFAULT '0',
  `CANT13` int(11) DEFAULT '0',
  `OBJ14` int(11) DEFAULT '0',
  `CANT14` int(11) DEFAULT '0',
  `OBJ15` int(11) DEFAULT '0',
  `CANT15` int(11) DEFAULT '0',
  `OBJ16` int(11) DEFAULT '0',
  `CANT16` int(11) DEFAULT '0',
  `OBJ17` int(11) DEFAULT '0',
  `CANT17` int(11) DEFAULT '0',
  `OBJ18` int(11) DEFAULT '0',
  `CANT18` int(11) DEFAULT '0',
  `OBJ19` int(11) DEFAULT '0',
  `CANT19` int(11) DEFAULT '0',
  `OBJ20` int(11) DEFAULT '0',
  `CANT20` int(11) DEFAULT '0',
  `OBJ21` int(11) DEFAULT '0',
  `CANT21` int(11) DEFAULT '0',
  `OBJ22` int(11) DEFAULT '0',
  `CANT22` int(11) DEFAULT '0',
  `OBJ23` int(11) DEFAULT '0',
  `CANT23` int(11) DEFAULT '0',
  `OBJ24` int(11) DEFAULT '0',
  `CANT24` int(11) DEFAULT '0',
  `OBJ25` int(11) DEFAULT '0',
  `CANT25` int(11) DEFAULT '0',
  `OBJ26` int(11) DEFAULT '0',
  `CANT26` int(11) DEFAULT '0',
  `OBJ27` int(11) DEFAULT '0',
  `CANT27` int(11) DEFAULT '0',
  `OBJ28` int(11) DEFAULT '0',
  `CANT28` int(11) DEFAULT '0',
  `OBJ29` int(11) DEFAULT '0',
  `CANT29` int(11) DEFAULT '0',
  `OBJ30` int(11) DEFAULT '0',
  `CANT30` int(11) DEFAULT '0',
  `OBJ31` int(11) DEFAULT '0',
  `CANT31` int(11) DEFAULT '0',
  `OBJ32` int(11) DEFAULT '0',
  `CANT32` int(11) DEFAULT '0',
  `OBJ33` int(11) DEFAULT '0',
  `CANT33` int(11) DEFAULT '0',
  `OBJ34` int(11) DEFAULT '0',
  `CANT34` int(11) DEFAULT '0',
  `OBJ35` int(11) DEFAULT '0',
  `CANT35` int(11) DEFAULT '0',
  `OBJ36` int(11) DEFAULT '0',
  `CANT36` int(11) DEFAULT '0',
  `OBJ37` int(11) DEFAULT '0',
  `CANT37` int(11) DEFAULT '0',
  `OBJ38` int(11) DEFAULT '0',
  `CANT38` int(11) DEFAULT '0',
  `OBJ39` int(11) DEFAULT '0',
  `CANT39` int(11) DEFAULT '0',
  `OBJ40` int(11) DEFAULT '0',
  `CANT40` int(11) DEFAULT '0',
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charbanco`
--

LOCK TABLES `charbanco` WRITE;
/*!40000 ALTER TABLE `charbanco` DISABLE KEYS */;
INSERT INTO `charbanco` VALUES (944,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(945,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(946,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(947,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(948,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(949,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(950,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(951,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(952,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(953,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(954,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(955,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(956,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(957,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(958,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(959,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(960,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(962,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(963,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(964,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(965,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
/*!40000 ALTER TABLE `charbanco` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charcorreo`
--

DROP TABLE IF EXISTS `charcorreo`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charcorreo` (
  `Idmsj` int(11) NOT NULL AUTO_INCREMENT,
  `IndexPJ` int(11) DEFAULT '0',
  `Mensaje` varchar(100) CHARACTER SET latin1 NOT NULL DEFAULT '',
  `De` varchar(100) CHARACTER SET latin1 NOT NULL DEFAULT '',
  `Cantidad` int(11) DEFAULT '0',
  `Item` int(11) DEFAULT '0',
  PRIMARY KEY (`Idmsj`),
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM AUTO_INCREMENT=2932 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charcorreo`
--

LOCK TABLES `charcorreo` WRITE;
/*!40000 ALTER TABLE `charcorreo` DISABLE KEYS */;
INSERT INTO `charcorreo` VALUES (2922,944,'prueba','marius',0,0),(2926,0,'frfrgttgtgtgtgtgtgt','sdasd',0,0),(2928,962,'','',0,0),(2929,963,'','',0,0),(2930,964,'','',0,0),(2931,965,'','',0,0);
/*!40000 ALTER TABLE `charcorreo` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charfaccion`
--

DROP TABLE IF EXISTS `charfaccion`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charfaccion` (
  `IndexPJ` int(11) NOT NULL DEFAULT '0',
  `EjercitoReal` int(11) DEFAULT '0',
  `EjercitoCaos` int(11) DEFAULT '0',
  `EjercitoMili` int(11) DEFAULT '0',
  `Republicano` int(11) DEFAULT '0',
  `Ciudadano` int(11) DEFAULT '0',
  `Rango` int(11) DEFAULT '0',
  `Renegado` int(11) DEFAULT '0',
  `CiudMatados` int(11) DEFAULT '0',
  `ReneMatados` int(11) DEFAULT '0',
  `RepuMatados` int(11) DEFAULT '0',
  `CaosMatados` int(11) DEFAULT '0',
  `ArmadaMatados` int(11) DEFAULT '0',
  `MiliMatados` int(11) DEFAULT '0',
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charfaccion`
--

LOCK TABLES `charfaccion` WRITE;
/*!40000 ALTER TABLE `charfaccion` DISABLE KEYS */;
INSERT INTO `charfaccion` VALUES (944,0,0,0,0,0,0,1,0,0,0,0,1,0),(945,0,0,0,0,0,0,1,0,100,0,0,0,0),(946,0,0,0,1,0,0,0,0,0,0,0,0,0),(947,0,0,0,1,0,0,0,0,0,0,0,0,0),(948,0,0,0,1,0,0,0,0,0,0,0,0,0),(949,0,0,0,1,0,0,0,0,0,0,0,0,0),(950,0,0,0,1,0,0,0,0,0,0,0,0,0),(951,0,0,0,1,0,0,0,0,0,0,0,0,0),(952,0,0,0,1,0,0,0,0,0,0,0,0,0),(953,0,0,0,1,0,0,0,0,0,0,0,0,0),(954,0,0,0,1,0,0,0,0,0,0,0,0,0),(955,0,0,0,1,0,0,0,0,0,0,0,0,0),(956,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(957,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(958,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(959,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(960,0,0,0,0,1,0,0,0,0,0,0,0,0),(0,0,0,0,0,1,0,0,0,0,0,0,0,0),(962,0,0,0,0,1,0,0,0,0,0,0,0,0),(963,0,0,0,0,0,0,0,0,0,0,0,0,0),(964,0,0,0,0,0,0,0,0,0,0,0,0,0),(965,0,0,0,0,1,0,0,0,0,0,0,0,0);
/*!40000 ALTER TABLE `charfaccion` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charflags`
--

DROP TABLE IF EXISTS `charflags`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charflags` (
  `IndexPJ` int(11) NOT NULL AUTO_INCREMENT,
  `id` int(11) NOT NULL DEFAULT '0',
  `Nombre` varchar(100) NOT NULL DEFAULT '',
  `Ban` int(11) DEFAULT '0',
  `Navegando` int(11) DEFAULT '0',
  `Envenenado` int(11) DEFAULT '0',
  `Pena` int(11) DEFAULT '0',
  `Paralizado` int(11) DEFAULT '0',
  `Desnudo` int(11) DEFAULT '0',
  `Sed` int(11) DEFAULT '0',
  `Hambre` int(11) DEFAULT '0',
  `Escondido` int(11) DEFAULT '0',
  `Muerto` int(11) DEFAULT '0',
  `Online` tinyint(1) NOT NULL DEFAULT '0',
  `Torneos` int(10) unsigned NOT NULL DEFAULT '0',
  `Privilegio` tinyint(4) NOT NULL DEFAULT '0',
  `desc` varchar(50) NOT NULL DEFAULT '',
  PRIMARY KEY (`IndexPJ`),
  KEY `IndexPJ` (`IndexPJ`),
  KEY `id` (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=966 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charflags`
--

LOCK TABLES `charflags` WRITE;
/*!40000 ALTER TABLE `charflags` DISABLE KEYS */;
INSERT INTO `charflags` VALUES (944,304,'fedudok',0,0,0,0,0,0,0,0,0,0,0,0,11,''),(945,303,'marius',0,0,0,0,0,0,0,0,0,0,0,0,11,''),(946,303,'gfdasd',0,0,0,0,0,1,0,0,0,0,0,0,0,''),(961,304,'sdasd',0,0,0,0,0,0,0,0,0,0,0,0,0,''),(960,304,'testtres',0,0,0,0,0,1,1,1,0,0,0,0,0,''),(962,304,'gtgtgt',0,0,0,0,0,0,0,0,0,0,0,0,0,''),(963,304,'frfrfrgtgt',0,0,0,0,0,0,0,0,0,0,0,0,0,''),(964,304,'frfrfrgtgtgtgt',0,0,0,0,0,0,0,0,0,0,0,0,0,''),(965,304,'gtgtghhyhyhy',0,0,0,0,0,1,0,0,0,0,0,0,0,'');
/*!40000 ALTER TABLE `charflags` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charfrags`
--

DROP TABLE IF EXISTS `charfrags`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charfrags` (
  `IndexPJ` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `pj1` int(10) unsigned NOT NULL DEFAULT '0',
  `time1` int(11) NOT NULL DEFAULT '0',
  `pj2` int(10) unsigned NOT NULL DEFAULT '0',
  `time2` int(11) NOT NULL DEFAULT '0',
  `pj3` int(10) unsigned NOT NULL DEFAULT '0',
  `time3` int(11) NOT NULL DEFAULT '0',
  `pj4` int(10) unsigned NOT NULL DEFAULT '0',
  `time4` int(11) NOT NULL DEFAULT '0',
  `pj5` int(10) unsigned NOT NULL DEFAULT '0',
  `time5` int(11) NOT NULL DEFAULT '0',
  PRIMARY KEY (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charfrags`
--

LOCK TABLES `charfrags` WRITE;
/*!40000 ALTER TABLE `charfrags` DISABLE KEYS */;
/*!40000 ALTER TABLE `charfrags` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charguild`
--

DROP TABLE IF EXISTS `charguild`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charguild` (
  `IndexPJ` int(11) NOT NULL DEFAULT '0',
  `GuildIndex` int(11) NOT NULL DEFAULT '0',
  `AspiranteA` int(11) DEFAULT '0',
  `Miembro` varchar(50) NOT NULL DEFAULT '',
  `MotivoRechazo` varchar(200) NOT NULL DEFAULT '',
  `Pedidos` varchar(200) NOT NULL DEFAULT '',
  KEY `IndexPJ` (`IndexPJ`),
  KEY `GuildIndex` (`GuildIndex`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charguild`
--

LOCK TABLES `charguild` WRITE;
/*!40000 ALTER TABLE `charguild` DISABLE KEYS */;
INSERT INTO `charguild` VALUES (944,1,0,'','',''),(945,1,0,'','','Verga'),(946,0,0,'','',''),(947,0,0,'','',''),(948,0,0,'','',''),(949,0,0,'','',''),(950,0,0,'','',''),(951,0,0,'','',''),(952,0,0,'','',''),(953,0,0,'','',''),(954,0,0,'','',''),(955,0,0,'','',''),(957,0,0,'','',''),(958,0,0,'','',''),(959,0,0,'','',''),(960,0,0,'','',''),(0,0,0,'','',''),(962,0,0,'','',''),(963,0,0,'','',''),(964,0,0,'','',''),(965,0,0,'','','');
/*!40000 ALTER TABLE `charguild` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charhechizos`
--

DROP TABLE IF EXISTS `charhechizos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charhechizos` (
  `IndexPJ` int(11) NOT NULL DEFAULT '0',
  `H1` int(11) DEFAULT '0',
  `H2` int(11) DEFAULT '0',
  `H3` int(11) DEFAULT '0',
  `H4` int(11) DEFAULT '0',
  `H5` int(11) DEFAULT '0',
  `H6` int(11) DEFAULT '0',
  `H7` int(11) DEFAULT '0',
  `H8` int(11) DEFAULT '0',
  `H9` int(11) DEFAULT '0',
  `H10` int(11) DEFAULT '0',
  `H11` int(11) DEFAULT '0',
  `H12` int(11) DEFAULT '0',
  `H13` int(11) DEFAULT '0',
  `H14` int(11) DEFAULT '0',
  `H15` int(11) DEFAULT '0',
  `H16` int(11) DEFAULT '0',
  `H17` int(11) DEFAULT '0',
  `H18` int(11) DEFAULT '0',
  `H19` int(11) DEFAULT '0',
  `H20` int(11) DEFAULT '0',
  `H21` int(11) DEFAULT '0',
  `H22` int(11) DEFAULT '0',
  `H23` int(11) DEFAULT '0',
  `H24` int(11) DEFAULT '0',
  `H25` int(11) DEFAULT '0',
  `H26` int(11) DEFAULT '0',
  `H27` int(11) DEFAULT '0',
  `H28` int(11) DEFAULT '0',
  `H29` int(11) DEFAULT '0',
  `H30` int(11) DEFAULT '0',
  `H31` int(11) DEFAULT '0',
  `H32` int(11) DEFAULT '0',
  `H33` int(11) DEFAULT '0',
  `H34` int(11) DEFAULT '0',
  `H35` int(11) DEFAULT '0',
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charhechizos`
--

LOCK TABLES `charhechizos` WRITE;
/*!40000 ALTER TABLE `charhechizos` DISABLE KEYS */;
INSERT INTO `charhechizos` VALUES (944,2,59,53,52,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(945,2,59,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(946,2,59,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(947,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(948,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(949,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(950,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(951,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(952,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(953,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(954,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(955,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(957,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(958,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(959,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(960,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(962,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(963,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(964,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(965,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
/*!40000 ALTER TABLE `charhechizos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charinit`
--

DROP TABLE IF EXISTS `charinit`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charinit` (
  `IndexPJ` int(11) NOT NULL DEFAULT '0',
  `Genero` int(11) DEFAULT '0',
  `raza` int(11) DEFAULT '0',
  `Hogar` int(11) DEFAULT '0',
  `clase` int(11) DEFAULT '0',
  `Heading` int(11) DEFAULT '0',
  `Head` int(11) DEFAULT '0',
  `Body` int(11) DEFAULT '0',
  `Arma` int(11) DEFAULT '0',
  `Escudo` int(11) DEFAULT '0',
  `Casco` int(11) DEFAULT '0',
  `LastIP` varchar(300) NOT NULL DEFAULT '',
  `mapa` int(11) DEFAULT '0',
  `X` int(11) DEFAULT '0',
  `Y` int(11) DEFAULT '0',
  `PAREJA` varchar(300) NOT NULL DEFAULT '',
  `pj1` int(10) unsigned NOT NULL DEFAULT '0',
  `time1` int(11) NOT NULL DEFAULT '0',
  `pj2` int(10) unsigned NOT NULL DEFAULT '0',
  `time2` int(11) NOT NULL DEFAULT '0',
  `pj3` int(10) unsigned NOT NULL DEFAULT '0',
  `time3` int(11) NOT NULL DEFAULT '0',
  `pj4` int(10) unsigned NOT NULL DEFAULT '0',
  `time4` int(11) NOT NULL DEFAULT '0',
  `pj5` int(10) unsigned NOT NULL DEFAULT '0',
  `time5` int(11) NOT NULL DEFAULT '0',
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charinit`
--

LOCK TABLES `charinit` WRITE;
/*!40000 ALTER TABLE `charinit` DISABLE KEYS */;
INSERT INTO `charinit` VALUES (944,1,1,2,3,3,1,255,25,2,10,'127.0.0.1',248,49,52,'',0,0,0,0,0,0,0,0,0,0),(945,1,3,1,18,3,202,2,2,2,2,'127.0.0.1',248,50,54,'',0,0,0,0,0,0,0,0,0,0),(946,2,3,1,2,3,270,40,16,2,2,'127.0.0.1',184,46,25,'',0,0,0,0,0,0,0,0,0,0),(947,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(948,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(949,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(950,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(951,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(952,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(953,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(954,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(955,1,1,1,1,3,1,1,2,2,2,'127.0.0.1',185,50,78,'',0,0,0,0,0,0,0,0,0,0),(959,0,0,0,0,0,0,0,0,0,0,'',0,0,0,'',0,0,0,0,0,0,0,0,0,0),(960,1,1,0,3,3,1,21,10,20,1,'127.0.0.1',34,68,51,'',0,0,0,0,0,0,0,0,0,0),(0,2,1,0,3,3,70,351,1,2,2,'',34,22,60,'',0,0,0,0,0,0,0,0,0,0),(962,1,2,0,3,3,101,2,1,2,2,'',34,40,87,'',0,0,0,0,0,0,0,0,0,0),(963,0,0,0,0,0,0,0,0,0,0,'',0,0,0,'',0,0,0,0,0,0,0,0,0,0),(964,1,2,0,4,3,101,2,2,2,2,'',0,0,0,'',0,0,0,0,0,0,0,0,0,0),(965,2,2,0,4,3,170,39,2,2,2,'',37,57,77,'',0,0,0,0,0,0,0,0,0,0);
/*!40000 ALTER TABLE `charinit` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charinvent`
--

DROP TABLE IF EXISTS `charinvent`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charinvent` (
  `IndexPJ` int(11) DEFAULT '0',
  `OBJ1` int(11) DEFAULT '0',
  `CANT1` int(11) DEFAULT '0',
  `OBJ2` int(11) DEFAULT '0',
  `CANT2` int(11) DEFAULT '0',
  `OBJ3` int(11) DEFAULT '0',
  `CANT3` int(11) DEFAULT '0',
  `OBJ4` int(11) DEFAULT '0',
  `CANT4` int(11) DEFAULT '0',
  `OBJ5` int(11) DEFAULT '0',
  `CANT5` int(11) DEFAULT '0',
  `OBJ6` int(11) DEFAULT '0',
  `CANT6` int(11) DEFAULT '0',
  `OBJ7` int(11) DEFAULT '0',
  `CANT7` int(11) DEFAULT '0',
  `OBJ8` int(11) DEFAULT '0',
  `CANT8` int(11) DEFAULT '0',
  `OBJ9` int(11) DEFAULT '0',
  `CANT9` int(11) DEFAULT '0',
  `OBJ10` int(11) DEFAULT '0',
  `CANT10` int(11) DEFAULT '0',
  `OBJ11` int(11) DEFAULT '0',
  `CANT11` int(11) DEFAULT '0',
  `OBJ12` int(11) DEFAULT '0',
  `CANT12` int(11) DEFAULT '0',
  `OBJ13` int(11) DEFAULT '0',
  `CANT13` int(11) DEFAULT '0',
  `OBJ14` int(11) DEFAULT '0',
  `CANT14` int(11) DEFAULT '0',
  `OBJ15` int(11) DEFAULT '0',
  `CANT15` int(11) DEFAULT '0',
  `OBJ16` int(11) DEFAULT '0',
  `CANT16` int(11) DEFAULT '0',
  `OBJ17` int(11) DEFAULT '0',
  `CANT17` int(11) DEFAULT '0',
  `OBJ18` int(11) DEFAULT '0',
  `CANT18` int(11) DEFAULT '0',
  `OBJ19` int(11) DEFAULT '0',
  `CANT19` int(11) DEFAULT '0',
  `OBJ20` int(11) DEFAULT '0',
  `CANT20` int(11) DEFAULT '0',
  `OBJ21` int(11) DEFAULT '0',
  `CANT21` int(11) DEFAULT '0',
  `OBJ22` int(11) DEFAULT '0',
  `CANT22` int(11) DEFAULT '0',
  `OBJ23` int(11) DEFAULT '0',
  `CANT23` int(11) DEFAULT '0',
  `OBJ24` int(11) DEFAULT '0',
  `CANT24` int(11) DEFAULT '0',
  `OBJ25` int(11) DEFAULT '0',
  `CANT25` int(11) DEFAULT '0',
  `CASCOSLOT` int(11) DEFAULT '0',
  `ARMORSLOT` int(11) DEFAULT '0',
  `SHIELDSLOT` int(11) DEFAULT '0',
  `WEAPONSLOT` int(11) DEFAULT '0',
  `ANILLOSLOT` int(11) DEFAULT '0',
  `MUNICIONSLOT` int(11) DEFAULT '0',
  `BARCOSLOT` int(11) DEFAULT '0',
  `NUDISLOT` int(11) DEFAULT '0',
  `MONTUSLOT` int(11) DEFAULT '0',
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charinvent`
--

LOCK TABLES `charinvent` WRITE;
/*!40000 ALTER TABLE `charinvent` DISABLE KEYS */;
INSERT INTO `charinvent` VALUES (944,879,1,1601,1,1329,1,196,1,991,1,1345,1,389,1,387,15,743,1,388,7550,1139,3,1093,2,0,0,0,0,386,8475,391,2,402,1,1333,1,1252,1,872,1,1078,1,990,1,1358,1,1266,1,1276,2,21,12,0,17,0,0,0,0,0),(945,573,100,572,100,463,1,161,5,1147,1,879,1,1601,1,1546,1,1257,1,32,1,1342,1,26,14,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,10,0,0,0,0,0,0,0),(946,0,0,0,0,0,0,0,0,0,0,879,1,1601,1,1181,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,8,0,0,0,0,0),(947,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(948,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(949,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(950,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(951,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(952,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(953,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(954,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(955,573,100,572,100,574,1,463,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(959,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(960,879,1,1601,1,0,0,132,1,129,1,996,1,1346,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,6,5,0,0,0,0,0),(0,573,100,572,100,574,1,1283,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(962,573,100,572,100,574,1,464,1,461,100,879,1,1601,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(963,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(964,573,100,572,100,460,1,464,1,461,100,576,100,1601,1,879,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,4,0,3,0,0,0,0,0),(965,573,100,572,100,460,1,1284,1,461,97,576,100,1601,1,879,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
/*!40000 ALTER TABLE `charinvent` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charmascotafami`
--

DROP TABLE IF EXISTS `charmascotafami`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charmascotafami` (
  `IndexPJ` int(11) DEFAULT '0',
  `Tiene` int(11) DEFAULT '0',
  `Nombre` varchar(200) CHARACTER SET latin1 NOT NULL DEFAULT '',
  `Tipo` int(11) DEFAULT '0',
  `Level` int(11) DEFAULT '0',
  `ELU` int(11) DEFAULT '0',
  `Exp` int(11) DEFAULT '0',
  `MinHP` int(11) DEFAULT '0',
  `MaxHP` int(11) DEFAULT '0',
  `MinHIT` int(11) DEFAULT '0',
  `MaxHIT` int(11) DEFAULT '0',
  `MAS1` int(11) DEFAULT '0',
  `MAS2` int(11) DEFAULT '0',
  `MAS3` int(11) DEFAULT '0',
  `NroMascotas` int(11) DEFAULT '0',
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charmascotafami`
--

LOCK TABLES `charmascotafami` WRITE;
/*!40000 ALTER TABLE `charmascotafami` DISABLE KEYS */;
INSERT INTO `charmascotafami` VALUES (944,1,'asdasd',1,1,100,0,10,10,5,15,0,0,0,0),(945,1,'asdasdas',1,1,100,0,10,10,5,15,0,0,0,0),(946,1,'asdasd',3,1,100,0,10,10,7,20,0,0,0,0),(947,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(948,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(949,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(950,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(951,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(952,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(953,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(954,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(955,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(960,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(0,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(962,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(963,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(964,0,'',0,0,0,0,0,0,0,0,0,0,0,0),(965,0,'',0,0,0,0,0,0,0,0,0,0,0,0);
/*!40000 ALTER TABLE `charmascotafami` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charskills`
--

DROP TABLE IF EXISTS `charskills`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charskills` (
  `IndexPJ` int(11) DEFAULT '0',
  `SK1` int(11) DEFAULT '0',
  `SK2` int(11) DEFAULT '0',
  `SK3` int(11) DEFAULT '0',
  `SK4` int(11) DEFAULT '0',
  `SK5` int(11) DEFAULT '0',
  `SK6` int(11) DEFAULT '0',
  `SK7` int(11) DEFAULT '0',
  `SK8` int(11) DEFAULT '0',
  `SK9` int(11) DEFAULT '0',
  `SK10` int(11) DEFAULT '0',
  `SK11` int(11) DEFAULT '0',
  `SK12` int(11) DEFAULT '0',
  `SK13` int(11) DEFAULT '0',
  `SK14` int(11) DEFAULT '0',
  `SK15` int(11) DEFAULT '0',
  `SK16` int(11) DEFAULT '0',
  `SK17` int(11) DEFAULT '0',
  `SK18` int(11) DEFAULT '0',
  `SK19` int(11) DEFAULT '0',
  `SK20` int(11) DEFAULT '0',
  `SK21` int(11) DEFAULT '0',
  `SK22` int(11) DEFAULT '0',
  `SK23` int(11) DEFAULT '0',
  `SK24` int(11) DEFAULT '0',
  `SK25` int(11) DEFAULT '0',
  `SK26` int(11) DEFAULT '0',
  `SK27` int(11) DEFAULT '0',
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charskills`
--

LOCK TABLES `charskills` WRITE;
/*!40000 ALTER TABLE `charskills` DISABLE KEYS */;
INSERT INTO `charskills` VALUES (944,37,33,12,0,0,0,2,13,0,37,0,0,0,0,12,0,100,0,0,0,0,100,0,0,0,7,100),(945,17,5,0,0,0,0,0,2,0,1,0,0,0,0,10,0,0,0,0,0,0,0,0,0,0,100,100),(946,10,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(947,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(948,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(949,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(950,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(951,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(952,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(953,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(954,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(955,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(960,100,24,0,0,0,0,8,0,0,0,0,0,0,0,15,0,0,0,0,0,0,0,0,0,0,80,100),(0,0,0,0,0,0,0,10,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(962,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,10,0,0,0,0,0,0,0,0,0,0),(963,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(964,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0),(965,3,0,10,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
/*!40000 ALTER TABLE `charskills` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `charstats`
--

DROP TABLE IF EXISTS `charstats`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `charstats` (
  `IndexPJ` int(11) DEFAULT '0',
  `GLD` int(11) DEFAULT '0',
  `BANCO` int(11) DEFAULT '0',
  `MaxHP` int(11) DEFAULT '0',
  `MinHP` int(11) DEFAULT '0',
  `MaxMAN` int(11) DEFAULT '0',
  `MinMAN` int(11) DEFAULT '0',
  `MinSTA` int(11) DEFAULT '0',
  `MaxSTA` int(11) DEFAULT '0',
  `MaxHIT` int(11) DEFAULT '0',
  `MinHIT` int(11) DEFAULT '0',
  `MinAGU` int(11) DEFAULT '0',
  `MinHAM` int(11) DEFAULT '0',
  `SkillPtsLibres` int(11) DEFAULT '0',
  `VecesMurioUsuario` int(11) DEFAULT '0',
  `Exp` int(11) DEFAULT '0',
  `ELV` int(11) DEFAULT '0',
  `NpcsMuertes` int(11) DEFAULT '0',
  `ELU` int(11) DEFAULT '0',
  KEY `IndexPJ` (`IndexPJ`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `charstats`
--

LOCK TABLES `charstats` WRITE;
/*!40000 ALTER TABLE `charstats` DISABLE KEYS */;
INSERT INTO `charstats` VALUES (944,102502657,0,262,262,3240,3240,866,866,61,60,90,100,3,6,0,50,15,0),(945,17508,0,36,36,594,594,180,180,12,11,100,75,45,10,0,50,12,0),(946,2000000,0,42,42,1920,48,31,586,41,40,100,100,200,1,23838464,40,0,41743574),(947,0,0,17,17,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(948,0,0,19,19,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(949,0,0,16,16,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(950,0,0,17,17,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(951,0,0,19,19,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(952,0,0,19,19,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(953,0,0,16,16,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(954,0,0,18,18,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(955,0,0,18,18,50,50,40,40,2,1,100,100,0,0,0,1,0,300),(960,1791119,179,123,123,0,0,0,720,119,119,0,0,45,0,200000809,45,19,249712682),(0,10,0,17,17,0,0,40,40,2,1,80,80,0,0,0,1,0,400),(962,10,0,17,17,0,0,40,40,2,1,100,100,0,0,0,1,0,400),(963,10,0,17,17,0,0,40,40,2,1,100,100,0,0,0,1,0,400),(964,10,0,17,17,50,50,40,40,2,1,100,100,0,0,0,1,0,400),(965,10,0,17,17,50,50,0,40,2,1,80,80,0,5,54,1,0,400);
/*!40000 ALTER TABLE `charstats` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `cuentas`
--

DROP TABLE IF EXISTS `cuentas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cuentas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `nombre` varchar(60) NOT NULL,
  `Password` varchar(64) NOT NULL,
  `mail` varchar(100) NOT NULL,
  `ban` int(11) DEFAULT '0',
  `Online` tinyint(1) NOT NULL DEFAULT '0',
  `numpjs` int(11) DEFAULT '0',
  `Donador` int(10) unsigned NOT NULL DEFAULT '0',
  `creditos` int(11) NOT NULL DEFAULT '0',
  `betatest` tinyint(1) NOT NULL DEFAULT '0',
  `verificada` varchar(20) NOT NULL DEFAULT '',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=305 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `cuentas`
--

LOCK TABLES `cuentas` WRITE;
/*!40000 ALTER TABLE `cuentas` DISABLE KEYS */;
INSERT INTO `cuentas` VALUES (303,'a','1','a',0,0,2,1,0,0,''),(304,'b','1','a',0,0,7,0,0,0,'1');
/*!40000 ALTER TABLE `cuentas` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `extras`
--

DROP TABLE IF EXISTS `extras`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `extras` (
  `id` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `nombre` varchar(255) DEFAULT '',
  `valor` varchar(100) DEFAULT '',
  PRIMARY KEY (`id`),
  KEY `id` (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=14 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `extras`
--

LOCK TABLES `extras` WRITE;
/*!40000 ALTER TABLE `extras` DISABLE KEYS */;
/*!40000 ALTER TABLE `extras` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `guildsinfo`
--

DROP TABLE IF EXISTS `guildsinfo`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `guildsinfo` (
  `GuildIndex` int(11) NOT NULL AUTO_INCREMENT,
  `Founder` varchar(100) DEFAULT '',
  `GuildName` varchar(100) DEFAULT '',
  `Date` datetime DEFAULT NULL,
  `Antifaccion` varchar(100) DEFAULT '',
  `Alineacion` int(11) DEFAULT '0',
  `Leader` varchar(100) DEFAULT '',
  `Codexs` varchar(100) DEFAULT '',
  `Desc` varchar(100) DEFAULT '',
  `URL` varchar(100) DEFAULT '',
  `Noticia` varchar(100) DEFAULT '',
  PRIMARY KEY (`GuildIndex`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `guildsinfo`
--

LOCK TABLES `guildsinfo` WRITE;
/*!40000 ALTER TABLE `guildsinfo` DISABLE KEYS */;
INSERT INTO `guildsinfo` VALUES (1,'fedudok','asfrfrf','2017-09-17 22:09:16','0',4,'fedudok','asdada|asdada|asdada|asdada|||||','asdada','',''),(2,'fedudok','asdasd','2017-09-17 22:09:19','0',4,'fedudok','','','',''),(3,'fedudok','grfrffrfr','2017-09-17 22:09:20','0',4,'fedudok','','','',''),(4,'fedudok','fasdasdasd','2017-09-17 22:09:15','0',4,'fedudok','','','',''),(5,'fedudok','asdasdasd','2017-09-17 22:09:29','0',4,'fedudok','','','','');
/*!40000 ALTER TABLE `guildsinfo` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `guildsolicitudes`
--

DROP TABLE IF EXISTS `guildsolicitudes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `guildsolicitudes` (
  `IndexSol` int(11) NOT NULL AUTO_INCREMENT,
  `GuildIndex` int(11) DEFAULT '0',
  `IndexPJ` int(11) DEFAULT '0',
  `Nombre` varchar(100) DEFAULT '',
  `solicitud` varchar(100) DEFAULT '',
  PRIMARY KEY (`IndexSol`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `guildsolicitudes`
--

LOCK TABLES `guildsolicitudes` WRITE;
/*!40000 ALTER TABLE `guildsolicitudes` DISABLE KEYS */;
/*!40000 ALTER TABLE `guildsolicitudes` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2017-09-18  0:10:20
