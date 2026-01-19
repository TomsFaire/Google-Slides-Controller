module.exports = function (self) {
	self.setActionDefinitions({
		open_presentation: {
			name: 'Open Presentation',
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
	})
}
