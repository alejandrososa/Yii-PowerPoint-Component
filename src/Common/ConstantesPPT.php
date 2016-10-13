<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 21/09/2016
 * Hora: 16:34
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

/**
 * Interface ConstantesPPT
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
interface ConstantesPPT
{
    /* Properties file */
    const PPT_CREATOR                           = 'PHPOffice';
    const PPT_TITLE                             = 'Sample Title';
    const PPT_SUBJECT                           = 'Sample Subject';
    const PPT_DESCRIPTION                       = 'Sample Description';

    /* Alignment styles of Text */
    const TEXT_ALIGN_HORIZONTAL_GENERAL         = 'l';
    const TEXT_ALIGN_HORIZONTAL_LEFT            = 'l';
    const TEXT_ALIGN_HORIZONTAL_RIGHT           = 'r';
    const TEXT_ALIGN_HORIZONTAL_CENTER          = 'ctr';
    const TEXT_ALIGN_HORIZONTAL_JUSTIFY         = 'just';
    const TEXT_ALIGN_HORIZONTAL_DISTRIBUTED     = 'dist';
    const TEXT_ALIGN_VERTICAL_BASE              = 'base';
    const TEXT_ALIGN_VERTICAL_AUTO              = 'auto';
    const TEXT_ALIGN_VERTICAL_BOTTOM            = 'b';
    const TEXT_ALIGN_VERTICAL_TOP               = 't';
    const TEXT_ALIGN_VERTICAL_CENTER            = 'ctr';
    const TEXT_HEIGHT                           = 300;
    const TEXT_WIDTH                            = 600;
    const TEXT_OFFSET_X                         = 170;
    const TEXT_OFFSET_Y                         = 180;
    const TEXT_SIZE                             = 14;
    const TEXT_BREAK                            = '<br>';
    const TEXT_FORMAT_PERCENTE                  = '#%';

    /* Chart Types */
    const CHART_HEIGHT                          = 234;
    const CHART_WIDTH                           = 450;
    const CHART_OFFSET_X                        = 41;
    const CHART_OFFSET_Y                        = 18;
    const CHART_TYPE_AREA                       = 'area';
    const CHART_TYPE_BAR                        = 'bar';
    const CHART_TYPE_BAR_HORIZONTAL             = 'barHorizontal';
    const CHART_TYPE_BAR_STACKED                = 'barStacked';
    const CHART_TYPE_BAR_PERCENT_STACKED        = 'barPercentStacked';
    const CHART_TYPE_BAR_3D                     = 'bar3D';
    const CHART_TYPE_LINE                       = 'line';
    const CHART_TYPE_PIE                        = 'pie';
    const CHART_TYPE_PIE_3D                     = 'pie3D';
    const CHART_TYPE_SCATTER                    = 'scatter';

    /* Colors */
    const COLOR_BLACK                           = 'FF000000';
    const COLOR_WHITE                           = 'FFFFFFFF';
    const COLOR_RED                             = 'FFF44336';
    const COLOR_PINK                            = 'FFE91E63';
    const COLOR_PURPLE                          = 'FF9C27B0';
    const COLOR_DEEP_PURPLE                     = 'FF673AB7';
    const COLOR_INDIGO                          = 'FF3F51B5';
    const COLOR_BLUE                            = 'FF2196F3';
    const COLOR_LIGHT_BLUE                      = 'FF03A9F4';
    const COLOR_CYAN                            = 'FF00BCD4';
    const COLOR_TEAL                            = 'FF009688';
    const COLOR_GREEN                           = 'FF4CAF50';
    const COLOR_LIGHT_GREEN                     = 'FF8BC34A';
    const COLOR_LIME                            = 'FFCDDC39';
    const COLOR_YELLOW                          = 'FFFFEB3B';
    const COLOR_AMBER                           = 'FFFFC107';
    const COLOR_ORANGE                          = 'FFFF9800';
    const COLOR_DEEP_ORANGE                     = 'FFFF5722';
    const COLOR_GREY                            = 'FF9E9E9E';
    const COLOR_BLUE_GREY                       = 'FF607D8B';
    const COLOR_LIGHT_GREY                      = 'FFCFD8DC';
    const COLOR_DIVIDER                         = 'FFBDBDBD';
    const COLOR_PRIMARY_TEXT                    = 'FF212121';
    const COLOR_SECONDARY_TEXT                  = 'FF757575';

    //SYMBOLS
    const SYMBOL_CIRCLE                         = 'circle';
    const SYMBOL_DASH                           = 'dash';
    const SYMBOL_DIAMOND                        = 'diamond';
    const SYMBOL_DOT                            = 'dot';
    const SYMBOL_NONE                           = 'none';
    const SYMBOL_PLUS                           = 'plus';
    const SYMBOL_SQUARE                         = 'square';
    const SYMBOL_STAR                           = 'star';
    const SYMBOL_TRIANGLE                       = 'triangle';
    const SYMBOL_X                              = 'x';

    //MARGINS
    const DEFAULT_MARGIN_LEFT                   = 50;
    const DEFAULT_MARGIN_TOP                    = 100;
    const DEFAULT_SLIDE_WITH                    = 1040;
    const DEFAULT_SLIDE_HEIGTH                  = 720;

    //BOOLEANS
    const TRUE                                  = true;
    const FALSE                                 = false;
}