<?php
/**
 * Squiz_Sniffs_CSS_ClassDefinitionNameSpacingSniff.
 *
 * PHP version 5
 *
 * @category  PHP
 * @package   PHP_CodeSniffer
 * @author    Greg Sherwood <gsherwood@squiz.net>
 * @copyright 2006-2014 Squiz Pty Ltd (ABN 77 084 670 600)
 * @license   https://github.com/squizlabs/PHP_CodeSniffer/blob/master/licence.txt BSD Licence
 * @link      http://pear.php.net/package/PHP_CodeSniffer
 */

/**
 * Squiz_Sniffs_CSS_ClassDefinitionNameSpacingSniff.
 *
 * Ensure there are no blank lines between the names of classes/IDs.
 *
 * @category  PHP
 * @package   PHP_CodeSniffer
 * @author    Greg Sherwood <gsherwood@squiz.net>
 * @copyright 2006-2014 Squiz Pty Ltd (ABN 77 084 670 600)
 * @license   https://github.com/squizlabs/PHP_CodeSniffer/blob/master/licence.txt BSD Licence
 * @version   Release: @package_version@
 * @link      http://pear.php.net/package/PHP_CodeSniffer
 */
class Squiz_Sniffs_CSS_ClassDefinitionNameSpacingSniff implements PHP_CodeSniffer_Sniff
{

    /**
     * A list of tokenizers this sniff supports.
     */
    public array $supportedTokenizers = array('CSS');


    /**
     * Returns the token types that this sniff is interested in.
     *
     * @return int[]
     */
    public function register()
    {
        return array(T_OPEN_CURLY_BRACKET);

    }//end register()


    /**
     * Processes the tokens that this sniff is interested in.
     *
     * @param PHP_CodeSniffer_File $phpcsFile The file where the token was found.
     * @param int                  $stackPtr  The position in the stack where
     *                                        the token was found.
     *
     * @return void
     */
    public function process(PHP_CodeSniffer_File $phpcsFile, $stackPtr)
    {
        $tokens = $phpcsFile->getTokens();

        // Do not check nested style definitions as, for example, in @media style rules.
        $nested = $phpcsFile->findNext(T_OPEN_CURLY_BRACKET, ($stackPtr + 1), $tokens[$stackPtr]['bracket_closer']);
        if ($nested !== false) {
            return;
        }

        // Find the first blank line before this opening brace, unless we get
        // to another style definition, comment or the start of the file.
        $endTokens  = array(
                       T_OPEN_CURLY_BRACKET  => T_OPEN_CURLY_BRACKET,
                       T_CLOSE_CURLY_BRACKET => T_CLOSE_CURLY_BRACKET,
                       T_OPEN_TAG            => T_OPEN_TAG,
                      );
        $endTokens += PHP_CodeSniffer_Tokens::$commentTokens;

        $prev = $phpcsFile->findPrevious(PHP_CodeSniffer_Tokens::$emptyTokens, ($stackPtr - 1), null, true);

        $foundContent = false;
        $currentLine  = $tokens[$prev]['line'];
        for ($i = ($stackPtr - 1); $i >= 0; $i--) {
            if (isset($endTokens[$tokens[$i]['code']])) {
                break;
            }

            if ($tokens[$i]['line'] === $currentLine) {
                if ($tokens[$i]['code'] !== T_WHITESPACE) {
                    $foundContent = true;
                }

                continue;
            }

            // We changed lines.
            if (!$foundContent) {
                // Before we throw an error, make sure we are not looking
                // at a gap before the style definition.
                $prev = $phpcsFile->findPrevious(T_WHITESPACE, $i, null, true);
                if ($prev !== false
                    && !isset($endTokens[$tokens[$prev]['code']])
                ) {
                    $error = 'Blank lines are not allowed between class names';
                    $phpcsFile->addError($error, ($i + 1), 'BlankLinesFound');
                }

                break;
            }

            $foundContent = false;
            $currentLine  = $tokens[$i]['line'];
        }//end for

    }//end process()


}//end class
