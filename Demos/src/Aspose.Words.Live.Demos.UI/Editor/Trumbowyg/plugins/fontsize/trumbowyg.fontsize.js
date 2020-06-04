// 12/07/2019 by Dmitry Zotov

(function ($) {
	'use strict';

	$.extend(true, $.trumbowyg, {
		langs: {
			// jshint camelcase:false
			en: {
				fontsize: 'Font size',
				fontsizes: {
					'x-small': 'Extra small',
					'small': 'Small',
					'medium': 'Regular',
					'large': 'Large',
					'x-large': 'Extra large',
					'custom': 'Custom'
				},
				fontCustomSize: {
					title: 'Custom Font Size',
					label: 'Font Size',
					value: '48'
				}
			},
			es: {
				fontsize: 'Tamaño de Fuente',
				fontsizes: {
					'x-small': 'Extra pequeña',
					'small': 'Pegueña',
					'medium': 'Regular',
					'large': 'Grande',
					'x-large': 'Extra Grande',
					'custom': 'Customizada'
				},
				fontCustomSize: {
					title: 'Tamaño de Fuente Customizada',
					label: 'Tamaño de Fuente',
					value: '48'
				}
			},
			da: {
				fontsize: 'Skriftstørrelse',
				fontsizes: {
					'x-small': 'Ekstra lille',
					'small': 'Lille',
					'medium': 'Normal',
					'large': 'Stor',
					'x-large': 'Ekstra stor',
					'custom': 'Brugerdefineret'
				}
			},
			fr: {
				fontsize: 'Taille de la police',
				fontsizes: {
					'x-small': 'Très petit',
					'small': 'Petit',
					'medium': 'Normal',
					'large': 'Grand',
					'x-large': 'Très grand',
					'custom': 'Taille personnalisée'
				},
				fontCustomSize: {
					title: 'Taille de police personnalisée',
					label: 'Taille de la police',
					value: '48'
				}
			},
			de: {
				fontsize: 'Font size',
				fontsizes: {
					'x-small': 'Sehr klein',
					'small': 'Klein',
					'medium': 'Normal',
					'large': 'Groß',
					'x-large': 'Sehr groß',
					'custom': 'Benutzerdefiniert'
				},
				fontCustomSize: {
					title: 'Benutzerdefinierte Schriftgröße',
					label: 'Schriftgröße',
					value: '48'
				}
			},
			nl: {
				fontsize: 'Lettergrootte',
				fontsizes: {
					'x-small': 'Extra klein',
					'small': 'Klein',
					'medium': 'Normaal',
					'large': 'Groot',
					'x-large': 'Extra groot',
					'custom': 'Tilpasset'
				}
			},
			tr: {
				fontsize: 'Yazı Boyutu',
				fontsizes: {
					'x-small': 'Çok Küçük',
					'small': 'Küçük',
					'medium': 'Normal',
					'large': 'Büyük',
					'x-large': 'Çok Büyük',
					'custom': 'Görenek'
				}
			},
			zh_tw: {
				fontsize: '字體大小',
				fontsizes: {
					'x-small': '最小',
					'small': '小',
					'medium': '中',
					'large': '大',
					'x-large': '最大',
					'custom': '自訂大小',
				},
				fontCustomSize: {
					title: '自訂義字體大小',
					label: '字體大小',
					value: '48'
				}
			},
			pt_br: {
				fontsize: 'Tamanho da fonte',
				fontsizes: {
					'x-small': 'Extra pequeno',
					'small': 'Pequeno',
					'regular': 'Médio',
					'large': 'Grande',
					'x-large': 'Extra grande',
					'custom': 'Personalizado'
				},
				fontCustomSize: {
					title: 'Tamanho de Fonte Personalizado',
					label: 'Tamanho de Fonte',
					value: '48'
				}
			},
			it: {
				fontsize: 'Dimensioni del testo',
				fontsizes: {
					'x-small': 'Molto piccolo',
					'small': 'piccolo',
					'regular': 'normale',
					'large': 'grande',
					'x-large': 'Molto grande',
					'custom': 'Personalizzato'
				},
				fontCustomSize: {
					title: 'Dimensioni del testo personalizzato',
					label: 'Dimensioni del testo',
					value: '48'
				}
			},
			ko: {
				fontsize: '글꼴 크기',
				fontsizes: {
					'x-small': '아주 작게',
					'small': '작게',
					'medium': '보통',
					'large': '크게',
					'x-large': '아주 크게',
					'custom': '사용자 지정'
				},
				fontCustomSize: {
					title: '사용자 지정 글꼴 크기',
					label: '글꼴 크기',
					value: '48'
				}
			},
		}
	});
	// jshint camelcase:true

	var defaultOptions = {
		sizeList: [
			'x-small',
			'small',
			'medium',
			'large',
			'x-large'
		],
		allowCustomSize: true
	};

	// Add dropdown with font sizes
	$.extend(true, $.trumbowyg, {
		plugins: {
			fontsize: {
				init: function (trumbowyg) {
					trumbowyg.o.plugins.fontsize = $.extend({},
						defaultOptions,
						trumbowyg.o.plugins.fontsize || {}
					);

					trumbowyg.addBtnDef('fontsize', {
						dropdown: buildDropdown(trumbowyg)
					});
				}
			}
		}
	});

	function setFontSize(elements, size) {
		$.each(elements, function (index, el) {
			try {
				if (el.nodeType !== Node.TEXT_NODE) {
					setFontSize(el.childNodes, size);
				} else {
					var style = el.parentElement.style;
					style.fontSize = size;
					style.lineHeight = 'normal';
				}	
			} catch (e) {

			}
		});
	}

	function splitElement(element, container, size, index, changeLeftPart) {
		if (element === undefined)
			return;
		try {
			var prev = container.previousSibling;
			var offset = 0;
			try {
				while (prev !== null) { // need to find the offset in the parent element
					offset += prev.textContent.length;
					prev = prev.previousSibling;
				}
			} catch (e) { }
			index += offset;

			var text = element.innerText;
			if (!changeLeftPart && index === 0 || changeLeftPart && index === text.length) // nothing to split
			{
				setFontSize([element], size);
				return;
			}

			var first = text.substring(0, index);
			var last = text.substring(index);
			var clone = element.cloneNode(true);
			element.innerText = last;
			clone.innerText = first;
			
			element.parentNode.insertBefore(clone, element);
			if (changeLeftPart)
				setFontSize([clone], size);
			else
				setFontSize([element], size);
		} catch (e) { }
	}

	function getElements(range, array, element) {
		element.childNodes.forEach(function(item) {
			try {
				if (item.nodeType !== Node.TEXT_NODE)
					getElements(range, array, item);
				else if (range.intersectsNode(item)) {
					if (!array.includes(item.parentNode) && item.parentNode.nodeName === "SPAN")
						array.push(item.parentNode);
				} else
					return;
			} catch (e) {

			}
		});
	}

	function changeFontSize(trumbowyg, size) {
		trumbowyg.saveRange(); // set range from selection
		var s = size + 'pt';

		/*var text = trumbowyg.range.startContainer.parentElement;
		var selectedText = trumbowyg.getRangeText();
		if ($(text).html() === selectedText) {
			$(text).css('font-size', s);
		} else {*/

		// split first and last into two, and change elements between
		var r = trumbowyg.range;
		var between = [];
		getElements(r, between, r.commonAncestorContainer);
		
		var first = between.shift();
		var last = between.pop();

		splitElement(first, r.startContainer, s, r.startOffset, false);
		setFontSize(between, s);
		splitElement(last, r.endContainer, s, r.endOffset, true);

		trumbowyg.saveRange();
	}

	function buildDropdown(trumbowyg) {
		var dropdown = [];

		$.each(trumbowyg.o.plugins.fontsize.sizeList, function (index, size) {
			trumbowyg.addBtnDef('fontsize_' + size, {
				text: '<span style="font-size: ' + size + 'pt;">' + (trumbowyg.lang.fontsizes[size] || size) + '</span>',
				hasIcon: false,
				fn: function () {
					//trumbowyg.execCmd('fontSize', index + 1, true); // old

					changeFontSize(trumbowyg, size);
					return true;
				}
			});
			dropdown.push('fontsize_' + size);
		});

		if (trumbowyg.o.plugins.fontsize.allowCustomSize) {
			var customSizeButtonName = 'fontsize_custom';
			var customSizeBtnDef = {
				fn: function () {
					trumbowyg.openModalInsert(trumbowyg.lang.fontCustomSize.title,
						{
							size: {
								label: trumbowyg.lang.fontCustomSize.label,
								value: trumbowyg.lang.fontCustomSize.value
							}
						},
						function (values) {
							changeFontSize(trumbowyg, values.size);

							/* old
							 var text = trumbowyg.range.startContainer.parentElement;
							var selectedText = trumbowyg.getRangeText();
							if ($(text).html() === selectedText) {
								$(text).css('font-size', values.size);
							} else {
								trumbowyg.range.deleteContents();
								var html = '<span style="font-size: ' + values.size + 'pt;">' + selectedText + '</span>';
								var node = $(html)[0];
								trumbowyg.range.insertNode(node);
							}
							trumbowyg.saveRange();
							*/
							return true;
						}
					);
				},
				text: '<span style="font-size: medium;">' + trumbowyg.lang.fontsizes.custom + '</span>',
				hasIcon: false
			};
			trumbowyg.addBtnDef(customSizeButtonName, customSizeBtnDef);
			dropdown.push(customSizeButtonName);
		}

		return dropdown;
	}
})(jQuery);
