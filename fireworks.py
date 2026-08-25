import pygame
import random
import math

pygame.init()

WIDTH = 1400
HEIGHT = 900

screen = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("Fireworks Festival")

clock = pygame.time.Clock()

FPS = 60

WIND = random.uniform(-0.01, 0.01)

# ===================================
# PARTICLE
# ===================================

class Particle:

    def __init__(self, x, y, color):

        self.x = x
        self.y = y

        angle = random.uniform(0, 2 * math.pi)
        speed = random.uniform(2, 10)

        self.vx = math.cos(angle) * speed
        self.vy = math.sin(angle) * speed

        self.life = random.randint(60, 120)
        self.max_life = self.life

        self.color = color

        self.size = random.randint(2, 5)

    def update(self):

        self.vx += WIND
        self.vy += 0.05

        self.x += self.vx
        self.y += self.vy

        self.life -= 1

    def alive(self):
        return self.life > 0

    def draw(self, surface):

        alpha = int(255 * (self.life / self.max_life))

        glow = pygame.Surface((30, 30), pygame.SRCALPHA)

        pygame.draw.circle(
            glow,
            (*self.color, alpha),
            (15, 15),
            self.size
        )

        surface.blit(glow, (self.x - 15, self.y - 15))


# ===================================
# HEART PARTICLE
# ===================================

class HeartParticle(Particle):

    def __init__(self, x, y, color, vx, vy):
        super().__init__(x, y, color)

        self.vx = vx
        self.vy = vy


# ===================================
# ROCKET
# ===================================

class Rocket:

    def __init__(self, target_x=None, target_y=None):

        self.x = target_x if target_x else random.randint(
            100,
            WIDTH - 100
        )

        self.y = HEIGHT + random.randint(0, 100)

        self.target_y = target_y if target_y else random.randint(
            100,
            HEIGHT // 2
        )

        # b?n xi?n
        self.vx = random.uniform(-4, 4)

        self.vy = random.uniform(-16, -10)

        self.color = (
            random.randint(100, 255),
            random.randint(100, 255),
            random.randint(100, 255)
        )

        self.trail = []

        self.type = random.choice([
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "normal",
            "heart",
            "heart",
            "sw30a0",
            "swcCAMRAD",
            "swcFEBFCWDBA",
            "swcFS_ACT",


        ])

    def update(self):

        self.trail.append((self.x, self.y))

        if len(self.trail) > 20:
            self.trail.pop(0)

        self.x += self.vx
        self.y += self.vy

        self.vy += 0.04

        return self.y > self.target_y

    def draw(self, surface):

        for i, pos in enumerate(self.trail):

            alpha = int(
                150 * (i + 1) / len(self.trail)
            )

            trail = pygame.Surface(
                (12, 12),
                pygame.SRCALPHA
            )

            pygame.draw.circle(
                trail,
                (*self.color, alpha),
                (6, 6),
                3
            )

            surface.blit(
                trail,
                (pos[0] - 6, pos[1] - 6)
            )

        pygame.draw.circle(
            surface,
            self.color,
            (int(self.x), int(self.y)),
            4
        )

    def explode(self):

        particles = []

        if self.type == "heart":

            for i in range(220):

                t = (i / 220) * 2 * math.pi

                x = 16 * math.sin(t) ** 3

                y = -(
                    13 * math.cos(t)
                    - 5 * math.cos(2 * t)
                    - 2 * math.cos(3 * t)
                    - math.cos(4 * t)
                )

                particles.append(
                    HeartParticle(
                        self.x,
                        self.y,
                        self.color,
                        x * 0.4,
                        y * 0.4
                    )
                )

        elif self.type == "sw30a0":
            font = pygame.font.SysFont(
                "Arial",
                50,
                bold=True
            )

            text_surface = font.render(
                "SW30A0",
                True,
                (255, 255, 255)
            )

            mask = pygame.mask.from_surface(
                text_surface
            )

            for px in range(mask.get_size()[0]):

                for py in range(mask.get_size()[1]):

                    if mask.get_at((px, py)):

                        p = Particle(
                            self.x,
                            self.y,
                            self.color
                        )

                        p.vx = (
                            px - mask.get_size()[0] / 2
                        ) * 0.08

                        p.vy = (
                            py - mask.get_size()[1] / 2
                        ) * 0.08

                        particles.append(p)

        elif self.type == "swcCAMRAD":
                    font = pygame.font.SysFont(
                        "Arial",
                        50,
                        bold=True
                    )
        
                    text_surface = font.render(
                        "swcCAM_RAD",
                        True,
                        (255, 255, 255)
                    )
        
                    mask = pygame.mask.from_surface(
                        text_surface
                    )
        
                    for px in range(mask.get_size()[0]):
        
                        for py in range(mask.get_size()[1]):
        
                            if mask.get_at((px, py)):
        
                                p = Particle(
                                    self.x,
                                    self.y,
                                    self.color
                                )
        
                                p.vx = (
                                    px - mask.get_size()[0] / 2
                                ) * 0.08
        
                                p.vy = (
                                    py - mask.get_size()[1] / 2
                                ) * 0.08
        
                                particles.append(p)


        elif self.type == "swcFEBFCWDBA":
                    font = pygame.font.SysFont(
                        "Arial",
                        50,
                        bold=True
                    )
        
                    text_surface = font.render(
                        "swcFEBFCWDBA",
                        True,
                        (255, 255, 255)
                    )
        
                    mask = pygame.mask.from_surface(
                        text_surface
                    )
        
                    for px in range(mask.get_size()[0]):
        
                        for py in range(mask.get_size()[1]):
        
                            if mask.get_at((px, py)):
        
                                p = Particle(
                                    self.x,
                                    self.y,
                                    self.color
                                )
        
                                p.vx = (
                                    px - mask.get_size()[0] / 2
                                ) * 0.08
        
                                p.vy = (
                                    py - mask.get_size()[1] / 2
                                ) * 0.08
        
                                particles.append(p)

        elif self.type == "swcFS_ACT":
                            font = pygame.font.SysFont(
                                "Arial",
                                50,
                                bold=True
                            )
                
                            text_surface = font.render(
                                "swcFS_ACT",
                                True,
                                (255, 255, 255)
                            )
                
                            mask = pygame.mask.from_surface(
                                text_surface
                            )
                
                            for px in range(mask.get_size()[0]):
                
                                for py in range(mask.get_size()[1]):
                
                                    if mask.get_at((px, py)):
                
                                        p = Particle(
                                            self.x,
                                            self.y,
                                            self.color
                                        )
                
                                        p.vx = (
                                            px - mask.get_size()[0] / 2
                                        ) * 0.08
                
                                        p.vy = (
                                            py - mask.get_size()[1] / 2
                                        ) * 0.08
                
                                        particles.append(p)

        elif self.type == "n_common_lib":
                                    font = pygame.font.SysFont(
                                        "Arial",
                                        50,
                                        bold=True
                                    )
                        
                                    text_surface = font.render(
                                        "n_common_lib",
                                        True,
                                        (255, 255, 255)
                                    )
                        
                                    mask = pygame.mask.from_surface(
                                        text_surface
                                    )
                        
                                    for px in range(mask.get_size()[0]):
                        
                                        for py in range(mask.get_size()[1]):
                        
                                            if mask.get_at((px, py)):
                        
                                                p = Particle(
                                                    self.x,
                                                    self.y,
                                                    self.color
                                                )
                        
                                                p.vx = (
                                                    px - mask.get_size()[0] / 2
                                                ) * 0.08
                        
                                                p.vy = (
                                                    py - mask.get_size()[1] / 2
                                                ) * 0.08
                        
                                                particles.append(p)

        elif self.type == "CCB22968":
                                            font = pygame.font.SysFont(
                                                "Arial",
                                                50,
                                                bold=True
                                            )
                                
                                            text_surface = font.render(
                                                "CCB22968",
                                                True,
                                                (255, 255, 255)
                                            )
                                
                                            mask = pygame.mask.from_surface(
                                                text_surface
                                            )
                                
                                            for px in range(mask.get_size()[0]):
                                
                                                for py in range(mask.get_size()[1]):
                                
                                                    if mask.get_at((px, py)):
                                
                                                        p = Particle(
                                                            self.x,
                                                            self.y,
                                                            self.color
                                                        )
                                
                                                        p.vx = (
                                                            px - mask.get_size()[0] / 2
                                                        ) * 0.08
                                
                                                        p.vy = (
                                                            py - mask.get_size()[1] / 2
                                                        ) * 0.08
                                
                                                        particles.append(p)

        else:

            count = random.randint(120, 300)

            for _ in range(count):
                particles.append(
                    Particle(
                        self.x,
                        self.y,
                        self.color
                    )
                )

        return particles


# ===================================
# STAR BACKGROUND
# ===================================

stars = []

for _ in range(300):
    stars.append(
        (
            random.randint(0, WIDTH),
            random.randint(0, HEIGHT),
            random.randint(1, 3)
        )
    )

# ===================================
# MAIN
# ===================================

rockets = []
particles = []

auto_timer = 0
finale_timer = 0

running = True

while running:

    dt = clock.tick(FPS)

    auto_timer += dt
    finale_timer += dt

    for event in pygame.event.get():

        if event.type == pygame.QUIT:
            running = False

        elif event.type == pygame.MOUSEBUTTONDOWN:

            mx, my = pygame.mouse.get_pos()

            rockets.append(
                Rocket(
                    target_x=mx,
                    target_y=my
                )
            )

    # t? ??ng b?n

    if auto_timer > 700:

        for _ in range(random.randint(1, 3)):
            rockets.append(Rocket())

        auto_timer = 0

    # finale

    if finale_timer > 30000:

        for _ in range(40):
            rockets.append(Rocket())

        finale_timer = 0

    screen.fill((5, 5, 20))

    # sao

    for s in stars:

        pygame.draw.circle(
            screen,
            (255, 255, 255),
            (s[0], s[1]),
            s[2]
        )

    # rocket

    temp_rockets = []

    for rocket in rockets:

        if rocket.update():
            temp_rockets.append(rocket)
        else:
            particles.extend(
                rocket.explode()
            )

        rocket.draw(screen)

    rockets = temp_rockets

    # particle

    alive = []

    for p in particles:

        p.update()

        if p.alive():
            alive.append(p)
            p.draw(screen)

    particles = alive

    pygame.display.flip()

pygame.quit()
