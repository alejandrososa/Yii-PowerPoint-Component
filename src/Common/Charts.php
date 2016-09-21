<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 20/09/2016
 * Hora: 15:14
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Style\Border;
use AlejandroSosa\YiiPowerPoint\Common\Helper;

/**
 * Class Charts
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Charts extends AbstractObject
{
    const PIE_HEIGHT        = 234;
    const PIE_WIDTH         = 450;
    const PIE_OFFSET_X      = 41;
    const PIE_OFFSET_Y      = 18;
    const PIE_BACKGROUND    = 'FFFFFFFF';

    /**
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function create(Slide $slide, $options = [])
    {
        // TODO: Implement create() method.
        return 'hola desde charts';
    }

    /**
     * Create chart Pie 3D
     * @param Slide $slide
     * @param string $title
     * @param array $series_data
     * @param array $options
     */
    public static function createPie3D(Slide $slide, $title = '', $series_data = [], $options = [])
    {
        $height         = Helper::hasArrayProperty('height', $options) ? $options['height'] : self::PIE_HEIGHT;
        $width          = Helper::hasArrayProperty('width', $options) ? $options['width'] : self::PIE_WIDTH;
        $offset_x       = Helper::hasArrayProperty('ox', $options) ? $options['ox'] : self::PIE_OFFSET_X;
        $offset_y       = Helper::hasArrayProperty('oy', $options) ? $options['oy'] : self::PIE_OFFSET_Y;
        $background     = Helper::hasArrayProperty('background', $options) ? $options['background'] : self::PIE_BACKGROUND;

        if($slide instanceof Slide) {
            $oFill = new Fill();
            $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($background));
            $oShadow = new Shadow();
            $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

            // Create a pie chart (that should be inserted in a shape)
            $pie3DChart = new Pie3D();
            $pie3DChart->setExplosion(20);
            $series = new Series('Downloads', $series_data);
            $series->setLabelPosition(Series::LABEL_BESTFIT);
            $series->setShowSeriesName(false);
            $series->setShowPercentage(true);
            $pie3DChart->addSeries($series);
            // Create a shape (chart)
            $shape = $slide->createChartShape();
            $shape->setName($title)
                ->setResizeProportional(false)
                ->setHeight($height)->setWidth($width)
                ->setOffsetX($offset_x)->setOffsetY($offset_y);
//            $shape->setShadow($oShadow);
            $shape->setFill($oFill);
            $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
            $shape->getTitle()->setText($title);
            $shape->getTitle()->getFont()->setItalic(true);
            $shape->getPlotArea()->setType($pie3DChart);
            $shape->getView3D()->setRotationX(30);
            $shape->getView3D()->setPerspective(30);
            $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
            $shape->getLegend()->getFont()->setItalic(true);
        }
    }
}