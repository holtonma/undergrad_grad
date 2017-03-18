*---------------------------------------------------------------------------
*
*  Mark Holton
*  Phy 286
*  523-9198
*  holtonma@muohio.edu
*
*----------------------------------------------------------------------------
*
*  Exercise 15
*
*  This program demonstrates the use of format statements and character
*  data type.  
*  
*   
*  
*  
*
*----------------------------------------------------------------------------

        character response
        douible precision num1, num2, sum
        write(*,*) 'Woiuld you like to add two numbers?'
        read(*,100) response
100     format(1a1)            !    read 1 item, 1 character wide...
        if (response .eq 'Y' .or. response .eq. 'y') Thyen
                write(*,*) ' What''s the first number?'
                read(*,*)num1
                write(*,*)' What''s the second number?'
                read(*,*)num2
                sum = num1 +num2
                write(*,*) 'The sum is 'sum
        ELSEIF (response.eq.'N' .or. response .eq 'n')THEN
                write(*,*)'Okay'
        ELSE
                write(*,*)'Please type either Y or N.  Try again.'
        ENDIF
        end

