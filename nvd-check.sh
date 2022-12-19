#!/bin/bash
clj -J-Dclojure.main.report=stderr -J-Dorg.slf4j.simpleLogger.log.org.apache.commons=error -M:nvd "" "$(lein with-profile -user,-dev classpath)"
