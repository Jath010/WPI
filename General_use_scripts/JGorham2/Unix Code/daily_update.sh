#!/bin/sh
#
# Rudimentary Requirements
#
if [ `whoami` != "root" ];
then
  echo "daily_update: root?"
  exit 1
fi

# 
# Setup variables for directories 
LBIN=/usr/local/bin
LSBIN=/usr/local/sbin
BASE=/tools/Security/PostOracle
architecture=`uname -s`
BASEBIN=${BASE}/${architecture}-`uname -m`
PROCESSED=/shared/elwood/processed
export ORACLE_SERVICE=/shared/oracle_service/unix
UGUYS=/home/uguys
PRINC="ldap-manager/`hostname | tr A-Z a-z`"

#
# change to the oracle report export area.
#
if cd ${ORACLE_SERVICE};
then
  continue
else
  echo "daily_update: ${ORACLE_SERVICE} wasn't available."
  exit 1
fi

#
# Look for departed employees
#
# First, were any prior runs forgotten?
#
if [ `ls xstaff.* 2>/dev/null | wc -l` -gt 0 ];
then
  echo "****"
  echo "**** NOTE: there are prior staff deletions which have not been run!"
  echo "****"
  echo "**** `ls xstaff.*`"
  echo "****"
fi
#
# Look for newly departed
#
if [ `ls employee_changes.* 2> /dev/null | wc -l` -gt 0 ];
then
  stamp=`date +%Y%m%d`
  rm -f /var/tmp/DEL.${stamp}
  for E in employee_changes.*;
  do
#
# Fields in employee_changes are separated by ~  e.g.
#
# The first field is the PIDM and the second field be marked "DEL"
#
# 518649~DEL~F2~145628640~Liang~Haodong~~~~~~~~~Mathematical Sciences~hdliang@WPI.EDU
#
    awk -F~ '$2=="DEL" && $2!="RE" && $2!="EP" { print $1 }' ${E} >> /var/tmp/DEL.${stamp}
    awk -F~ '$2=="DEL" && $2~/(RE|EP)/ { print $1,$2 }' ${E} >> /var/tmp/RE.${stamp}
    rm -f ${E}
  done
  for PIDM in `cat /var/tmp/DEL.${stamp}`;
  do
    echo "select unix_login from person_by_pidm where unix_login is not null and pidm
= '$PIDM'" | ${BASE}/mysql --skip-column-names >> ${BASE}/xstaff.${stamp}
  done
  for PIDM in `cat /var/tmp/RE.${stamp}`;
  do
    echo "select unix_login from person_by_pidm where unix_login is not null and pidm
= '$PIDM'" | ${BASE}/mysql --skip-column-names >> ${BASE}/restaff.${stamp}
  done
  rm /var/tmp/DEL.${stamp} /var/tmp/RE.${stamp}
  if [ -s ${BASE}/xstaff.${stamp} ];
  then
    if [ `cat ${BASE}/xstaff.${stamp} | wc -l` -eq 1 ];
    then
      plurality="is a potential staff login deletion"
    else
      plurality="are potential staff login deletions"
    fi
    echo "daily_update: There ${plurality} in ${BASE}/xstaff.${stamp}"
    echo ""
  else
    rm -f ${BASE}/xstaff.${stamp}
  fi
  if [ -s ${BASE}/restaff.${stamp} ];
  then
    if [ `cat ${BASE}/restaff.${stamp} | wc -l` -eq 1 ];
    then
      plurality="is a potential staff login retirement"
    else
      plurality="are potential staff login retirements"
    fi
    echo "daily_update: There ${plurality} in ${BASE}/restaff.${stamp}"
    echo ""
  else
    rm -f ${BASE}/restaff.${stamp}
  fi
fi

#
# Put the class lists in the traditional class list location while trimming the carriage returns from the data
# 
# The z95 file is used to identify advisors.
#
if [ -f z95.class_list ];
then
  if [ -s z95.class_list ];
  then
    any=/bin/true
    echo -n "${intro}Advisors"
    intro=" "
    sed -e 's/\r//g' < z95.class_list > ${UGUYS}/z95.classlists
    chown uguys.uguys ${UGUYS}/z95.classlists
    chmod 0660 ${UGUYS}/z95.classlists
  fi
  rm z95.class_list
fi
# The other classlists  are named for the terms which the files describe: <term-letter><two-digit-year>
#
# Loop through class list files, avoiding error of empty set
#
# Copy the lists from the source area to the UGUYS director.
# Trim the carriage return when copying.
#
# The data is used by various software:
#
# There's a classlist shell script: /usr/local/bin/classlist
#
# The project registrations in those files defines whether a student can submit a project
# and also it associates the partners in a project and the advisors of a project.
#

lists=`ls *.list_class 2>/dev/null | wc -l`

if [ ${lists} -gt 0 ];
then
  any=/bin/true
  echo -n "${intro}ClassList"
  intro=" "
  if [ ${lists} -gt 1 ];

  then
    echo -n "s"
  fi
  for list in *.list_class;
  do
    class=`basename ${list} .list_class | tr A-Z a-z`
    echo -n " ${class}"
    sed -e 's/\r//g' < ${list} > ${UGUYS}/${class}.classlists
    chown uguys.uguys ${UGUYS}/${class}.classlists
    chmod 0660 ${UGUYS}/${class}.classlists
    rm ${list}
  done
fi
#
if ${any};
then
  echo ""
fi

#
# set universal sort order
#
LC_ALL=C
export LC_ALL
#
# Update database with new info from banner
#
export FALSE=/bin/false
export TRUE=/bin/true
export CHOWN=/bin/chown
export LDAPSEARCH=${LBIN}/ldapsearch
export LDAPMODIFY=${LBIN}/ldapmodify
export MAIL=/bin/mail
export SENDMAIL=/usr/sbin/sendmail
export K5START=${LBIN}/k5start

#
# Move on to the student registration data.
#
student_files=`ls aej_student_????????.dat 2>/dev/null | wc -l`
if [ ${student_files} -gt 0 -o -f ${ORACLE_SERVICE}/banner_directory_extract.txt ];
then
#
# Find a probable Oracle version.
#
  try=/usr/local/oracle/product/11.1.0/db_1/
  tried=${try}
  if [ -d  $try ];
  then
    export ORACLE_HOME=${try}
  else
    echo "daily_update: No Oracle Home; tried ${tried}"
    exit 1
  fi
#
# clean up failed prior?
#
  rm -f student direct majors
#
  if [ ${student_files} -gt 0 ];
  then
#
# perhaps combine the files
#
    a=""
    b=""
    for subset in aej_student_????????.dat;
    do
      year=`echo ${subset} | awk '{print substr($1,13,4)}'`
      semester=`echo ${subset} | awk '{print substr($1,17,2)}'`
      direct=`echo ${subset} | awk '{print substr($1,19,1)}'`
      enrolled=`echo ${subset} | awk '{print substr($1,20,1)}'`
      cut=${year}${semester}${direct}${enrolled}
      echo "daily_update: Students as defined by ${year} ${semester} ${direct} ${enrolled}"
      if [ ! -s aej_direct_${cut}.dat -o ! -s aej_majors_${cut}.dat ];
      then
	echo "daily_update: The student, direct, and majors files of ${cut} do not all exist!"
	exit 1
      fi
      if [ "X${a}" = "X" ];
      then
	a=${cut}
      else
	if [ "X${b}" = "X" ];
	then
	  b=${cut}
	else
	  echo "daily_update: There were too many course list subsets (2 max)."
	  echo "daily_update: Found subsets ${a}, ${b}, ${cut}..."
	  exit 1
	fi
      fi
    done
    if [ "X${b}" = "X" ];
    then
      ln aej_student_${a}.dat student
      ln aej_direct_${a}.dat direct
      ln aej_majors_${a}.dat majors
    else
#
# ${BASE}/uniq_id.c, installed in ${LSBIN}/ will take a couple input files
# and write out a single output, combining the info on a given ID which
# might be in both files into one line.
#
      if ${LSBIN}/uniq_id aej_student_${a}.dat aej_student_${b}.dat > student;
      then
	continue
      else
	exit 1
      fi
      if ${LSBIN}/uniq_id aej_direct_${a}.dat aej_direct_${b}.dat > direct;
      then
	continue
      else
	exit 1
      fi
      if ${LSBIN}/uniq_id aej_majors_${a}.dat aej_majors_${b}.dat > majors;
      then
        continue
      else
	exit 1
      fi
    fi
  fi
#
# The banner_directory_extract is the employee file.
#
  if [ -f ${ORACLE_SERVICE}/banner_directory_extract.txt ];
  then
    echo "daily_update: banner_directory_extract.txt found"
#
# update_sql moves banner_directory_extract.txt from ORACLE_SERVICE to
# PROCESSED. Make a copy of the prior file, for comparison later on.
#
    if cp ${PROCESSED}/banner_directory_extract.txt ${PROCESSED}/banner_directory_extract.old;
    then
      continue
    else
      echo "daily_update: banner_directory_extract copy failed"
      exit 1
    fi
  fi
#
# ${BASE}/update_sql.c looks for registration and staff files
# and combines the data on an individual and updates the userbase database on the webdb
# mysql server.
#
# The userbase database is used by a multitude of processes
#
  if ${LSBIN}/update_sql
  then
    rm -f majors student direct
    if [ ${student_files} -gt 0 ];
    then
      for subset in aej_student_????????.dat;
      do
	cut=`echo ${subset} | awk '{print substr($1,13,8)}'`
	rm aej_direct_${cut}.dat aej_majors_${cut}.dat aej_student_${cut}.dat
      done
    fi
  else
    echo ""
    echo "daily_update: **** SQL Update Failing."
    echo ""
    exit 1
  fi
else
#
# If neither student nor staff data, there's no reason to run an update.
#
  cat <<EOF

daily_update: There is no new student or staff data

EOF
  exit 0
fi
#
# ${BASE}/orphan-alum.c collects all the PIDMs from the alumni extract
# and removes alum status from an people who are no longer in that extract.
#
${BASEBIN}/orphan-alum
#
# ${BASE}/alumnify.c looks through the alum extract and sets
# the alum column in the person_by_pidm table of the userbase.
#
${LSBIN}/alumnify > /dev/null
#
# ${BASE}/verify_sql.c reads the recent staff and registration
# files and updates the person_by_pidm table with the date.  If an scan is run
# to remove "old" logins, the date controls whether the login is removed or not.
#
${LSBIN}/verify_sql > /dev/null 2>&1
cd ${BASE}
#
# /tools/Security/PostOracle/src/ldap-create.c loads the ldap info and the userbase
# database info and will update an existing LDAP record with new info, or create
# an LDAP entry for a newly appeared person.  -v sets 'chatty', giving runtime.
# -r specifies recipient(s) for "new departments" mailing.
#
#04/10/2019 - temporary disable.
#####${LBIN}/ldap-create -v -r lapierre
#
# Report on personnel changes.  The ${BASE}/check_and_balance.c program
# reads in the last staff list and the current one and compares them.  If there is a change
# which concerns the listed employee codes, it emails the recipient list with a report of
# the changes.
#
# This came into being when the Provost's Office was surprised by the assignments that
# Human Resources was making.  The most pressing problem was when HR was terminating
# Emeritus professors (employee code EP) unexpectedly, which used to result in removal
# of login.
#
if [ -f ${PROCESSED}/banner_directory_extract.old ];
then
  if ${LSBIN}/check_and_balance ${PROCESSED}/banner_directory_extract.old ${PROCESSED}/banner_directory_extract.txt -emp_code FT FN FA F1 F2 DH EP NE -recipient djgraves@wpi.edu kecoffey@wpi.edu costello@wpi.edu dcroke@wpi.edu soppong@wpi.edu lapierre@wpi.edu cjmararian@wpi.edu;
  then
    rm ${PROCESSED}/banner_directory_extract.old
  else
#
# if the program fails for some reason, don't remove the comparison file, since it really hasn't been processed.
#
    echo "daily_update: Did not remove ${PROCESSED}/banner_directory_extract.old due to check_and_balance failure."
  fi
fi
#
# Remove all but the newest generation of data in /shared/elwood/processed/
#
${BASE}/remove_redundant

#
# Update the advisor info in userbase
#
${BASE}/daily_advise

