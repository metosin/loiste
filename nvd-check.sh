#!/bin/bash
lein run -m nvd.task.check "" "$(lein with-profile -user,-dev classpath)"
