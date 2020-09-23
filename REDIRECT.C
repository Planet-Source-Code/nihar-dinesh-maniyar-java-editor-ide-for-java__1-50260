#include <sys\stat.h>
#include <string.h>
#include <fcntl.h>
#include <io.h>
#include<process.h>
#include<stdio.h>
int main(int argc,char *argv[])
{
   #define STDOUT 1
   #define STDERR 2
   int nul, oldstderr,oldstdout;

   char cmdLine[200];

   int argvLen=strlen(argv[0]);
   while( argv[0][--argvLen] != '\\');

   argv[0][argvLen+1]='\0';


   strcpy(cmdLine,argv[0]);
   strcat(cmdLine,"OUTPUT");
   /*create a file*/
   nul = open(cmdLine,O_CREAT | O_TRUNC | O_WRONLY | O_TEXT,S_IWRITE);

   /* create a duplicate handle for standard output */
   oldstdout = dup(STDOUT);
   oldstderr = dup(STDERR);

   /*
      redirect standard output to DUMMY.FIL
      by duplicating the file handle onto
      the file handle for standard output.
   */
   dup2(nul, STDOUT);
   dup2(nul, STDERR);

   /* close the handle for DUMMY.FIL */
   close(nul);

   /* will be redirected into DUMMY.FIL */

   strcpy(cmdLine,"\"");
   strcat(cmdLine,argv[0]);
   strcat(cmdLine,"editor.bat\"");
   system(cmdLine);
   /*write(STDOUT, msg, strlen(msg));

    restore original standard output handle */
   dup2(oldstderr, STDERR);
   dup2(oldstdout, STDOUT);

   /* close duplicate handle for STDOUT */
   close(oldstderr);
   close(oldstdout);
   return 0;
}
