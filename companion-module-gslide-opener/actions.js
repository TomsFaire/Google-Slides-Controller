module.exports = function (self) {
	console.log('[gslide-opener] actions.js - Setting up action definitions')
	self.log('info', '=== Setting up action definitions ===')
	
	const actionDefinitions = {
		open_presentation: {
			name: 'Open Presentation',
			description: 'Open a presentation (without auto-starting speaker notes)',
			options: [
				{
					id: 'url',
					type: 'textinput',
					label: 'Google Slides URL',
					default: '',
					required: true,
					useVariables: true,
				},
			],
			callback: async (event) => {
				try {
					const url = await self.parseVariablesInString(event.options.url)
					self.log('info', `Opening presentation: ${url}`)

					const response = await self.apiRequest('POST', '/api/open-presentation', { url })
					self.log('info', response.message || 'Presentation opened')
				} catch (error) {
					self.log('error', `Failed to open presentation: ${error.message}`)
				}
			},
		},

		open_presentation_with_notes: {
			name: 'Open Presentation with Notes',
			description: 'Open a presentation and automatically start speaker notes',
			options: [
				{
					id: 'url',
					type: 'textinput',
					label: 'Google Slides URL',
					default: '',
					required: true,
					useVariables: true,
				},
			],
			callback: async (event) => {
				try {
					const url = await self.parseVariablesInString(event.options.url)
					self.log('info', `Opening presentation with notes: ${url}`)

					const response = await self.apiRequest('POST', '/api/open-presentation-with-notes', { url })
					self.log('info', response.message || 'Presentation opened with notes')
				} catch (error) {
					self.log('error', `Failed to open presentation with notes: ${error.message}`)
				}
			},
		},

		open_preset_1: {
			name: 'Open Presentation 1',
			description: 'Open the preset presentation configured as "Presentation 1"',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Opening preset presentation 1')
					const response = await self.apiRequest('POST', '/api/open-preset', { preset: 1 })
					self.log('info', response.message || 'Preset 1 opened')
				} catch (error) {
					self.log('error', `Failed to open preset 1: ${error.message}`)
				}
			},
		},

		open_preset_2: {
			name: 'Open Presentation 2',
			description: 'Open the preset presentation configured as "Presentation 2"',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Opening preset presentation 2')
					const response = await self.apiRequest('POST', '/api/open-preset', { preset: 2 })
					self.log('info', response.message || 'Preset 2 opened')
				} catch (error) {
					self.log('error', `Failed to open preset 2: ${error.message}`)
				}
			},
		},

		open_preset_3: {
			name: 'Open Presentation 3',
			description: 'Open the preset presentation configured as "Presentation 3"',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Opening preset presentation 3')
					const response = await self.apiRequest('POST', '/api/open-preset', { preset: 3 })
					self.log('info', response.message || 'Preset 3 opened')
				} catch (error) {
					self.log('error', `Failed to open preset 3: ${error.message}`)
				}
			},
		},

		close_presentation: {
			name: 'Close Current Presentation',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Closing presentation')
					const response = await self.apiRequest('POST', '/api/close-presentation', {})
					self.log('info', response.message || 'Presentation closed')
				} catch (error) {
					self.log('error', `Failed to close presentation: ${error.message}`)
				}
			},
		},

		next_slide: {
			name: 'Next Slide',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Next slide')
					const response = await self.apiRequest('POST', '/api/next-slide', {})
					self.log('debug', response.message || 'Next slide')
				} catch (error) {
					self.log('error', `Failed to go to next slide: ${error.message}`)
				}
			},
		},

		previous_slide: {
			name: 'Previous Slide',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Previous slide')
					const response = await self.apiRequest('POST', '/api/previous-slide', {})
					self.log('debug', response.message || 'Previous slide')
				} catch (error) {
					self.log('error', `Failed to go to previous slide: ${error.message}`)
				}
			},
		},

		go_to_slide: {
			name: 'Go to Slide',
			options: [
				{
					id: 'slide',
					type: 'number',
					label: 'Slide Number',
					default: 1,
					min: 1,
					required: true,
					useVariables: true,
				},
			],
			callback: async (event) => {
				try {
					const slideStr = await self.parseVariablesInString(String(event.options.slide))
					const slide = parseInt(slideStr, 10)
					
					if (isNaN(slide) || slide < 1) {
						self.log('error', `Invalid slide number: ${slideStr}`)
						return
					}

					self.log('info', `Navigating to slide ${slide}`)
					const response = await self.apiRequest('POST', '/api/go-to-slide', { slide })
					self.log('info', response.message || `Navigated to slide ${slide}`)
				} catch (error) {
					self.log('error', `Failed to go to slide: ${error.message}`)
				}
			},
		},

		reload_presentation: {
			name: 'Reload Presentation',
			description: 'Closes and reopens the current presentation, returning to the same slide',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Reloading presentation...')
					const response = await self.apiRequest('POST', '/api/reload-presentation', {})
					self.log('info', response.message || 'Presentation reloaded')
				} catch (error) {
					self.log('error', `Failed to reload presentation: ${error.message}`)
				}
			},
		},

		toggle_video: {
			name: 'Toggle Video Playback',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Toggling video playback')
					const response = await self.apiRequest('POST', '/api/toggle-video', {})
					self.log('debug', response.message || 'Video toggled')
				} catch (error) {
					self.log('error', `Failed to toggle video: ${error.message}`)
				}
			},
		},

		open_speaker_notes: {
			name: 'Open Speaker Notes',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Toggling speaker notes')
					const response = await self.apiRequest('POST', '/api/open-speaker-notes', {})
					self.log('debug', response.message || 'Speaker notes toggled')
				} catch (error) {
					self.log('error', `Failed to open speaker notes: ${error.message}`)
				}
			},
		},

		close_speaker_notes: {
			name: 'Close Speaker Notes',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Closing speaker notes')
					const response = await self.apiRequest('POST', '/api/close-speaker-notes', {})
					self.log('debug', response.message || 'Speaker notes closed')
				} catch (error) {
					self.log('error', `Failed to close speaker notes: ${error.message}`)
				}
			},
		},

		scroll_notes_down: {
			name: 'Scroll Speaker Notes Down',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Scrolling speaker notes down')
					const response = await self.apiRequest('POST', '/api/scroll-notes-down', {})
					self.log('debug', response.message || 'Notes scrolled down')
				} catch (error) {
					self.log('error', `Failed to scroll notes: ${error.message}`)
				}
			},
		},

		scroll_notes_up: {
			name: 'Scroll Speaker Notes Up',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Scrolling speaker notes up')
					const response = await self.apiRequest('POST', '/api/scroll-notes-up', {})
					self.log('debug', response.message || 'Notes scrolled up')
				} catch (error) {
					self.log('error', `Failed to scroll notes: ${error.message}`)
				}
			},
		},

		zoom_in_notes: {
			name: 'Zoom In Speaker Notes',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Zooming in on speaker notes')
					const response = await self.apiRequest('POST', '/api/zoom-in-notes', {})
					self.log('debug', response.message || 'Zoomed in')
				} catch (error) {
					self.log('error', `Failed to zoom in on notes: ${error.message}`)
				}
			},
		},

		zoom_out_notes: {
			name: 'Zoom Out Speaker Notes',
			options: [],
			callback: async () => {
				try {
					self.log('info', 'Zooming out on speaker notes')
					const response = await self.apiRequest('POST', '/api/zoom-out-notes', {})
					self.log('debug', response.message || 'Zoomed out')
				} catch (error) {
					self.log('error', `Failed to zoom out on notes: ${error.message}`)
				}
			},
		},
	}
	
	console.log('[gslide-opener] actions.js - Registering', Object.keys(actionDefinitions).length, 'actions')
	self.log('info', `Registering ${Object.keys(actionDefinitions).length} actions`)
	self.setActionDefinitions(actionDefinitions)
	console.log('[gslide-opener] actions.js - Actions registered successfully')
	self.log('info', 'Actions registered successfully')
}
