<?php

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Slide\Background\Color as BackgroundColor;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color as StyleColor;

class SubtitleGenerator {

	const COMMENT_PREFIX = '#';

	/** @var PhpPresentation */
	private $presentation;
	private $lines = array();

	public function __construct() {
		$this->init();
	}

	private function init() {
		$this->presentation = new PhpPresentation();
	}

	private function loadBook() {
		$this->lines = file('input/book.txt');
	}

	private function generateDemoBook($numberOfSlides) {
		for ($i = 1; $i <= $numberOfSlides; $i++) {
			$this->lines[] = 'Slide ' . $i;
			$this->lines[] = '#Comment ' .$i . 'a';
			$this->lines[] = '#Comment ' .$i . 'b';
		}
	}

	public function run() {
		$this->loadBook();
		//$this->generateDemoBook(10);

		$slideValues = array();

		for ($i = 0; $i < count($this->lines); $i++) {
			$slideValue = array();
			$slideValue['comments'] = array();

			while (
				($i < count($this->lines))
				&& (substr($this->lines[$i], 0, strlen(self::COMMENT_PREFIX)) === self::COMMENT_PREFIX)
			) {
				$slideValue['comments'][] = substr($this->lines[$i], strlen(self::COMMENT_PREFIX));
				$i++;
			}

			if ($i < count($this->lines)) {
				$slideValue['text'] = $this->lines[$i];
			} else {
				$slideValue['text'] = '';
			}

			$slideValues[] = $slideValue;
		}

		foreach ($slideValues as $slideValue) {
			$this->addTextAndCommentsToPresentation($slideValue['text'], $slideValue['comments']);
		}

		$this->writeFiles('/output/slides');
	}

	private function setPropertiesInPresentation() {
		$this->presentation->getDocumentProperties()
			->setCreator('PHPOffice')
			->setLastModifiedBy('PHPPresentation Team')
			->setTitle('Sample 01 Title')
			->setSubject('Sample 01 Subject')
			->setDescription('Sample 01 Description')
			->setKeywords('office 2007 openxml libreoffice odt php')
			->setCategory('Sample Category');
	}

	private function addTextAndCommentsToPresentation($text = '', array $comments = array()) {
		$slide = $this->presentation->createSlide();

		// background
		$backgroundColor = new BackgroundColor();
		$backgroundColor->setColor(new StyleColor(StyleColor::COLOR_BLACK));
		$slide->setBackground($backgroundColor);

		// text
		$shape = $slide->createRichTextShape()
			->setHeight(300)
			->setWidth(960)
			->setOffsetX(0)
			->setOffsetY(360);
		$shape->getActiveParagraph()->getAlignment()
			->setHorizontal( Alignment::HORIZONTAL_CENTER );
		$textRun = $shape->createTextRun($text);
		$textRun->getFont()
			->setBold(true)
			->setSize(30)
			->setColor( new StyleColor( StyleColor::COLOR_WHITE ) );

		// comments
		$note = $slide->getNote();
		$richText=$note->createRichTextShape();
		foreach ($comments as $i => $comment) {
			if ($i > 0) {
				$richText->createParagraph();
			}
			$textRun = $richText->createTextRun($comment);
			$textRun->getFont()
				->setBold(false)
				->setSize(20)
				->setColor( new StyleColor( StyleColor::COLOR_BLACK ) );
		}
	}

	private function writeFiles($pathAndBasename) {
		$oWriterPPTX = IOFactory::createWriter($this->presentation, 'PowerPoint2007');
		$oWriterPPTX->save(__DIR__ . $pathAndBasename . '.pptx');

		$oWriterODP = IOFactory::createWriter($this->presentation, 'ODPresentation');
		$oWriterODP->save(__DIR__ . $pathAndBasename . '.odp');
	}
}

$subtitleGenerator = new SubtitleGenerator();
$subtitleGenerator->run();





?>